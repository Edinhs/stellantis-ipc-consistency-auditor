"""
Parsers de documentos de engenharia da Stellantis (otimizados para pacotes de 1.000+ páginas):
- VFs em formato Word (.docx) e PDF (.pdf) (consolidados de 1.000+ páginas ou individuais)
- Ignora automaticamente dezenas de tabelas distratoras (CAN Signals, DTCs, Estados, Hardware Pinout)
- Extrai com precisão cirúrgica apenas a seção Indication Management e a Configuration Parameters Table
- Planilha Section C - HMI Requirements / Indications (.xlsx)
- Planilha Section B - Software Requirements / Parameters Table (.xlsx)
"""

import re
from pathlib import Path
import docx
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import openpyxl
import fitz  # PyMuPDF

from app.models import (
    VFIndicationItem,
    VFParameterItem,
    ParsedVFDocument,
    SheetIndicationRow,
    SheetParameterRow,
)


VF_ID_REGEX = re.compile(r"(?<![A-Z0-9])(VF\d{2,4})(?![A-Z0-9])", re.IGNORECASE)
VF_FULL_CODE_REGEX = re.compile(r"(?<![A-Z0-9])(VF\d{2,4}_V\d+_[A-Z0-9]+)(?![A-Z0-9])", re.IGNORECASE)
# Captura padrões de indicações no meio de textos corridos/bullets do DOORS, ex: "HAND BRAKE.V1", "AIR BAG FAIL.V1"
PROSE_INDICATION_REGEX = re.compile(
    r"(?<![A-Z0-9_])([A-Z][A-Z0-9\s\-/]{2,48}?\.V\d+)(?![A-Z0-9_])"
)
NUMBERED_SUBSECTION_REGEX = re.compile(r"^\d+\.\d+(?:\.\d+)?\s+\S+")


def normalize_indication_name(name: str) -> str:
    """
    Normaliza o nome da indicação para comparar VF ('HAND BRAKE.V1') com a planilha Indications ('HAND BRAKE V1').
    Substitui '.V<num>' por ' V<num>', remove pontuação redundante nas bordas e colapsa espaços.
    """
    if not name:
        return ""
    cleaned = name.strip().upper()
    cleaned = re.sub(r"\.V(\d+)\b", r" V\1", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def normalize_param_name(name: str) -> str:
    """Normaliza o nome do parâmetro preservando letras/números/underscores em caixa alta para comparação."""
    if not name:
        return ""
    return re.sub(r"\s+", " ", name.strip()).upper()


def extract_vf_ids(text: str) -> list[str]:
    """Extrai códigos curtos de VF (ex: ['VF143']) a partir de textos como 'VF143_V1_PJ3U' ou 'VF143 - Hand Brake'."""
    if not text:
        return []
    matches = VF_ID_REGEX.findall(str(text))
    seen = []
    for m in matches:
        up = m.upper()
        if up not in seen:
            seen.append(up)
    return seen


def evaluate_ipc_supplier_eol_filter(component: str, implementation: str) -> tuple[bool, str]:
    """
    Verifica se o parâmetro da VF pertence ao escopo da planilha IPC Parameters Table:
    - Component deve conter 'IPC'
    - Implementation deve ser 'SUPPLIER EOL' (ex: 'SUPPLIER EOL', 'SUPPLIER EOL_2') e NÃO 'PROXI'.
    Retorna (is_in_scope, reason).
    """
    comp_up = (component or "").strip().upper()
    impl_up = (implementation or "").strip().upper()

    is_ipc = "IPC" in comp_up
    is_proxi = "PROXI" in impl_up
    is_supplier_eol = ("SUPPLIER" in impl_up and "EOL" in impl_up) or impl_up.startswith("EOL")

    if not is_ipc:
        return False, f"Fora do escopo IPC: Componente é '{component.strip()}' (pertence a outra ECU)."
    if is_proxi:
        return False, f"Fora do escopo Parameters Table EOL: Implementação é '{implementation.strip()}' (especificado no documento PROXI)."
    if not is_supplier_eol:
        return False, f"Fora do escopo SUPPLIER EOL: Implementação é '{implementation.strip()}'."

    return True, "Escopo Auditável (Component = IPC & Implementation = SUPPLIER EOL)"


def iter_block_items(parent):
    """Itera em ordem sequencial pelos parágrafos e tabelas de um documento Word (.docx)."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Elemento pai inválido para iter_block_items")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def parse_vf_docx(file_path: str | Path) -> list[ParsedVFDocument]:
    """
    Lê um arquivo .docx de VF (seja um documento consolidado com dezenas de VFs e 1.000+ páginas
    como VFS_IPC_PJ3U_2A.docx ou um arquivo individual de VF) e extrai cada VF, suas indicações
    e seus parâmetros de configuração, ignorando tabelas distratoras (CAN, DTC, Pinout, etc.).
    """
    path = Path(file_path)
    doc = docx.Document(str(path))
    vfs_by_id: dict[str, ParsedVFDocument] = {}
    current_vf: ParsedVFDocument | None = None
    current_section: str = ""

    file_vf_ids = extract_vf_ids(path.stem)
    default_vf_id = file_vf_ids[0] if file_vf_ids else None

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue

            style_name = (block.style.name if block.style else "").lower()
            vf_ids_in_text = extract_vf_ids(text)
            full_code_match = VF_FULL_CODE_REGEX.search(text)

            # Detectar início de uma nova VF em títulos (Heading 1) ou linhas numeradas de capítulo principal
            is_vf_heading = (
                ("heading 1" in style_name or re.match(r"^\d+\s+VF\d+", text, re.IGNORECASE))
                and vf_ids_in_text
            )
            if is_vf_heading:
                vf_id = vf_ids_in_text[0]
                full_code = full_code_match.group(1).upper() if full_code_match else vf_id
                clean_title = re.sub(r"^\d+\s+", "", text)
                clean_title = re.sub(r"\s*\(VF\d+[^)]*\)\s*$", "", clean_title).strip()
                if vf_id not in vfs_by_id:
                    vfs_by_id[vf_id] = ParsedVFDocument(
                        vf_id=vf_id,
                        vf_full_code=full_code,
                        title=clean_title,
                        source_file=path.name,
                    )
                current_vf = vfs_by_id[vf_id]
                current_section = ""
                continue

            if current_vf and full_code_match and "document code" in text.lower():
                current_vf.vf_full_code = full_code_match.group(1).upper()

            low_text = text.lower()
            if "indication management" in low_text:
                current_section = "INDICATIONS"
                continue
            elif "configuration parameters" in low_text:
                current_section = "PARAMETERS"
                continue
            elif style_name.startswith("heading") or NUMBERED_SUBSECTION_REGEX.match(text):
                # Se encontrou outro subtítulo numerado (ex: "62.13 Failure Management", "62.4 CAN Signals"), sai da seção atual
                if current_section == "INDICATIONS" and "indication" not in low_text:
                    current_section = ""
                elif current_section == "PARAMETERS" and "parameter" not in low_text:
                    current_section = ""

            # Extração de indicações em texto corrido/bullets DOORS dentro de Indication Management
            if current_vf and current_section == "INDICATIONS":
                existing_norms = {i.normalized_name for i in current_vf.indications}
                for match in PROSE_INDICATION_REGEX.finditer(text):
                    raw_cand = match.group(1).strip()
                    if VF_ID_REGEX.search(raw_cand):
                        continue
                    norm_cand = normalize_indication_name(raw_cand)
                    if norm_cand and norm_cand not in existing_norms:
                        existing_norms.add(norm_cand)
                        current_vf.indications.append(
                            VFIndicationItem(
                                vf_id=current_vf.vf_id,
                                vf_full_code=current_vf.vf_full_code,
                                vf_title=current_vf.title,
                                raw_name=raw_cand,
                                normalized_name=norm_cand,
                                condition=text[:140],
                                popup_message="Extraído do texto DOORS (Indication Management)",
                                telltale_behavior="-",
                                source_file=path.name,
                            )
                        )

        elif isinstance(block, Table):
            if not block.rows:
                continue

            if current_vf is None and default_vf_id:
                current_vf = vfs_by_id.setdefault(
                    default_vf_id,
                    ParsedVFDocument(
                        vf_id=default_vf_id,
                        vf_full_code=default_vf_id,
                        title=path.stem.replace("_", " "),
                        source_file=path.name,
                    ),
                )

            if current_vf is None:
                continue

            headers = [c.text.strip().lower() for c in block.rows[0].cells]
            header_joined = " | ".join(headers)

            # Ignorar tabelas distratoras explícitas (CAN Signals, DTCs, State Machines, Revisions)
            if any(distractor in header_joined for distractor in ["can message", "signal name", "dtc code", "pin number", "state transition"]):
                continue

            # Verificar se é tabela de Indicações
            is_ind_table = ("indication name" in header_joined) or (
                current_section == "INDICATIONS" and "parameter name" not in header_joined
            )
            if is_ind_table:
                for row in block.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if not cells or not cells[0] or cells[0] == "-":
                        continue
                    raw_ind_name = cells[0]
                    norm_ind = normalize_indication_name(raw_ind_name)
                    if not norm_ind or norm_ind in ("INDICATION NAME", "SIGNAL NAME", "STATE"):
                        continue
                    existing_item = next((x for x in current_vf.indications if x.normalized_name == norm_ind), None)
                    if existing_item:
                        existing_item.raw_name = raw_ind_name
                        existing_item.condition = cells[1] if len(cells) > 1 else existing_item.condition
                        existing_item.popup_message = cells[2] if len(cells) > 2 else existing_item.popup_message
                        existing_item.telltale_behavior = cells[3] if len(cells) > 3 else existing_item.telltale_behavior
                    else:
                        current_vf.indications.append(
                            VFIndicationItem(
                                vf_id=current_vf.vf_id,
                                vf_full_code=current_vf.vf_full_code,
                                vf_title=current_vf.title,
                                raw_name=raw_ind_name,
                                normalized_name=norm_ind,
                                condition=cells[1] if len(cells) > 1 else "",
                                popup_message=cells[2] if len(cells) > 2 else "",
                                telltale_behavior=cells[3] if len(cells) > 3 else "",
                                source_file=path.name,
                            )
                        )
                continue

            # Verificar se é tabela de Configuration Parameters
            is_param_table = (
                "parameter name" in header_joined
                or "first trial" in header_joined
                or (current_section == "PARAMETERS" and "implementation" in header_joined)
            )
            if is_param_table:
                def find_col(keywords: list[str], default_idx: int) -> int:
                    for idx, h in enumerate(headers):
                        if any(k in h for k in keywords):
                            return idx
                    return default_idx

                idx_id = find_col(["id"], 0)
                idx_name = find_col(["parameter name", "name"], 1)
                idx_ftv = find_col(["first trial"], 2)
                idx_range = find_col(["range"], 3)
                idx_res = find_col(["resolution"], 4)
                idx_unit = find_col(["unit"], 5)
                idx_comp = find_col(["component"], 6)
                idx_impl = find_col(["implementation"], 7)
                idx_xcp = find_col(["xcp"], 8)
                idx_impact = find_col(["impact"], 9)
                idx_desc = find_col(["description"], 10)

                for row in block.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) <= idx_name:
                        continue
                    p_name = cells[idx_name]
                    if not p_name or p_name in ("-", "None", "N/A", "Parameter Name"):
                        continue

                    comp_val = cells[idx_comp] if len(cells) > idx_comp else ""
                    impl_val = cells[idx_impl] if len(cells) > idx_impl else ""
                    is_in_scope, reason = evaluate_ipc_supplier_eol_filter(comp_val, impl_val)

                    current_vf.parameters.append(
                        VFParameterItem(
                            vf_id=current_vf.vf_id,
                            vf_full_code=current_vf.vf_full_code,
                            vf_title=current_vf.title,
                            param_id=cells[idx_id] if len(cells) > idx_id else "",
                            name=p_name,
                            normalized_name=normalize_param_name(p_name),
                            first_trial_value=cells[idx_ftv] if len(cells) > idx_ftv else "",
                            range_value=cells[idx_range] if len(cells) > idx_range else "",
                            resolution=cells[idx_res] if len(cells) > idx_res else "",
                            unit=cells[idx_unit] if len(cells) > idx_unit else "",
                            component=comp_val,
                            implementation=impl_val,
                            xcp_availability=cells[idx_xcp] if len(cells) > idx_xcp else "",
                            impact=cells[idx_impact] if len(cells) > idx_impact else "",
                            description=cells[idx_desc] if len(cells) > idx_desc else "",
                            is_ipc_supplier_eol=is_in_scope,
                            filter_reason=reason,
                            source_file=path.name,
                        )
                    )

    return list(vfs_by_id.values())


def parse_vf_pdf(file_path: str | Path) -> list[ParsedVFDocument]:
    """
    Lê um arquivo .pdf de VF (incluindo pacotes consolidados de 1.000+ páginas com dezenas de VFs e
    múltiplas páginas por VF) e extrai cada VF, suas indicações e parâmetros.
    """
    path = Path(file_path)
    pdf_doc = fitz.open(str(path))
    vfs_by_id: dict[str, ParsedVFDocument] = {}
    current_vf: ParsedVFDocument | None = None
    current_section: str = ""

    for page in pdf_doc:
        text = page.get_text("text")
        if not text:
            continue
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for ln in lines:
            # Ignorar rodapés/cabeçalhos repetitivos de página
            if ln.startswith(("STELLANTIS - EE SYSTEM", "Page ", "CONFIDENTIAL - STELLANTIS")):
                continue

            vf_ids = extract_vf_ids(ln)
            full_match = VF_FULL_CODE_REGEX.search(ln)
            if (
                vf_ids
                and ("-" in ln or full_match)
                and not ln.startswith(("IND_ROW", "PAR_ROW", "CAN_ROW", "DTC_ROW", "Document Code:", "[REQ_"))
                and re.match(r"^(?:\d+\s+)?VF\d+", ln, re.IGNORECASE)
            ):
                vf_id = vf_ids[0]
                full_code = full_match.group(1).upper() if full_match else vf_id
                clean_title = re.sub(r"^\d+\s+", "", ln)
                clean_title = re.sub(r"\s*\(VF\d+[^)]*\)\s*$", "", clean_title).strip()
                current_vf = vfs_by_id.setdefault(
                    vf_id,
                    ParsedVFDocument(
                        vf_id=vf_id,
                        vf_full_code=full_code,
                        title=clean_title,
                        source_file=path.name,
                    ),
                )
                current_section = ""
                continue

            if current_vf and ln.startswith("Document Code:") and full_match:
                current_vf.vf_full_code = full_match.group(1).upper()
                continue

            low_ln = ln.lower()
            if "indication management" in low_ln:
                current_section = "INDICATIONS"
                continue
            elif "configuration parameters" in low_ln:
                current_section = "PARAMETERS"
                continue
            elif NUMBERED_SUBSECTION_REGEX.match(ln):
                if current_section == "INDICATIONS" and "indication" not in low_ln:
                    current_section = ""
                elif current_section == "PARAMETERS" and "parameter" not in low_ln:
                    current_section = ""

            if current_vf and ln.startswith("IND_ROW |"):
                parts = [p.strip() for p in ln.split("|")[1:]]
                if parts and parts[0] and parts[0] != "-":
                    norm_ind = normalize_indication_name(parts[0])
                    if not any(x.normalized_name == norm_ind for x in current_vf.indications):
                        current_vf.indications.append(
                            VFIndicationItem(
                                vf_id=current_vf.vf_id,
                                vf_full_code=current_vf.vf_full_code,
                                vf_title=current_vf.title,
                                raw_name=parts[0],
                                normalized_name=norm_ind,
                                condition=parts[1] if len(parts) > 1 else "",
                                popup_message=parts[2] if len(parts) > 2 else "",
                                telltale_behavior=parts[3] if len(parts) > 3 else "",
                                source_file=path.name,
                            )
                        )
            elif current_vf and ln.startswith("PAR_ROW |"):
                parts = [p.strip() for p in ln.split("|")[1:]]
                if len(parts) >= 8 and parts[1] and parts[1] != "-":
                    comp_val = parts[6]
                    impl_val = parts[7]
                    is_in_scope, reason = evaluate_ipc_supplier_eol_filter(comp_val, impl_val)
                    current_vf.parameters.append(
                        VFParameterItem(
                            vf_id=current_vf.vf_id,
                            vf_full_code=current_vf.vf_full_code,
                            vf_title=current_vf.title,
                            param_id=parts[0],
                            name=parts[1],
                            normalized_name=normalize_param_name(parts[1]),
                            first_trial_value=parts[2],
                            range_value=parts[3],
                            resolution=parts[4],
                            unit=parts[5],
                            component=comp_val,
                            implementation=impl_val,
                            xcp_availability=parts[8] if len(parts) > 8 else "",
                            impact=parts[9] if len(parts) > 9 else "",
                            description=parts[10] if len(parts) > 10 else "",
                            is_ipc_supplier_eol=is_in_scope,
                            filter_reason=reason,
                            source_file=path.name,
                        )
                    )
            elif current_vf and current_section == "INDICATIONS":
                # Também suporta extração de texto corrido DOORS em PDFs dentro da seção Indication Management
                existing_norms = {i.normalized_name for i in current_vf.indications}
                for match in PROSE_INDICATION_REGEX.finditer(ln):
                    raw_cand = match.group(1).strip()
                    if VF_ID_REGEX.search(raw_cand):
                        continue
                    norm_cand = normalize_indication_name(raw_cand)
                    if norm_cand and norm_cand not in existing_norms:
                        existing_norms.add(norm_cand)
                        current_vf.indications.append(
                            VFIndicationItem(
                                vf_id=current_vf.vf_id,
                                vf_full_code=current_vf.vf_full_code,
                                vf_title=current_vf.title,
                                raw_name=raw_cand,
                                normalized_name=norm_cand,
                                condition=ln[:140],
                                popup_message="Extraído de texto PDF (Indication Management)",
                                telltale_behavior="-",
                                source_file=path.name,
                            )
                        )

    pdf_doc.close()
    return list(vfs_by_id.values())


def parse_indications_excel(file_path: str | Path) -> list[SheetIndicationRow]:
    """
    Lê a planilha Indications (ex: Indications_J3U_R4.xlsx).
    Prioriza a aba 'Full TFT 7 Inch. - LATAM' ou 'All IPC' (como visto no vídeo aos 10:15).
    """
    path = Path(file_path)
    wb = openpyxl.load_workbook(str(path), data_only=True)

    preferred_sheets = ["Full TFT 7 Inch. - LATAM", "All IPC"]
    selected_sheet = None
    for s_name in preferred_sheets:
        if s_name in wb.sheetnames:
            selected_sheet = wb[s_name]
            break

    if selected_sheet is None:
        for s_name in wb.sheetnames:
            ws = wb[s_name]
            first_row = [str(c.value or "").strip().lower() for c in ws[1]]
            if "name" in first_row and "vfs" in first_row:
                selected_sheet = ws
                break

    if selected_sheet is None:
        selected_sheet = wb.active

    headers = [str(c.value or "").strip().lower() for c in selected_sheet[1]]

    def col_idx(names: list[str], default: int) -> int:
        for i, h in enumerate(headers):
            if any(n == h or n in h for n in names):
                return i
        return default

    i_ipc = col_idx(["ipc"], 0)
    i_name = col_idx(["name"], 1)
    i_vfs = col_idx(["vfs", "vf"], 2)
    i_hard = col_idx(["hard tt code"], 3)
    i_soft = col_idx(["soft tt code"], 4)
    i_driven = col_idx(["driven"], 5)
    i_behav = col_idx(["hard tt std behav"], 6)
    i_msg = col_idx(["display messages"], 9)

    rows_out: list[SheetIndicationRow] = []
    for r_idx, row in enumerate(selected_sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not row or len(row) <= i_name:
            continue
        raw_name = str(row[i_name] or "").strip()
        if not raw_name or raw_name.lower() in ("none", "not applicable for high v10 audit"):
            continue
        vfs_raw = str(row[i_vfs] or "").strip() if len(row) > i_vfs and row[i_vfs] is not None else ""
        rows_out.append(
            SheetIndicationRow(
                sheet_name=selected_sheet.title,
                row_number=r_idx,
                ipc_variant=str(row[i_ipc] or "").strip() if len(row) > i_ipc and row[i_ipc] is not None else "",
                raw_name=raw_name,
                normalized_name=normalize_indication_name(raw_name),
                vfs_raw=vfs_raw,
                extracted_vf_ids=extract_vf_ids(vfs_raw),
                hard_tt_code=str(row[i_hard] or "").strip() if len(row) > i_hard and row[i_hard] is not None else "",
                soft_tt_code=str(row[i_soft] or "").strip() if len(row) > i_soft and row[i_soft] is not None else "",
                driven=str(row[i_driven] or "").strip() if len(row) > i_driven and row[i_driven] is not None else "",
                std_behav=str(row[i_behav] or "").strip() if len(row) > i_behav and row[i_behav] is not None else "",
                display_messages=str(row[i_msg] or "").strip() if len(row) > i_msg and row[i_msg] is not None else "",
                source_file=path.name,
            )
        )

    wb.close()
    return rows_out


def parse_parameters_excel(file_path: str | Path) -> list[SheetParameterRow]:
    """
    Lê a planilha IPC Parameters Table (ex: IPC_Parameters_Table_PJ3U_R5.xlsx).
    Prioriza a aba 'Parameters Table' (como visto no vídeo aos 19:15).
    """
    path = Path(file_path)
    wb = openpyxl.load_workbook(str(path), data_only=True)

    selected_sheet = wb["Parameters Table"] if "Parameters Table" in wb.sheetnames else None
    if selected_sheet is None:
        for s_name in wb.sheetnames:
            ws = wb[s_name]
            first_row = [str(c.value or "").strip().lower() for c in ws[1]]
            if "parameter name" in first_row:
                selected_sheet = ws
                break
    if selected_sheet is None:
        selected_sheet = wb.active

    headers = [str(c.value or "").strip().lower() for c in selected_sheet[1]]

    def col_idx(names: list[str], default: int) -> int:
        for i, h in enumerate(headers):
            if any(n == h or n in h for n in names):
                return i
        return default

    i_name = col_idx(["parameter name", "name"], 0)
    i_ftv = col_idx(["first trial value", "first trial"], 1)
    i_range = col_idx(["range value", "range"], 2)
    i_res = col_idx(["resolution"], 3)
    i_unit = col_idx(["unit"], 4)
    i_desc = col_idx(["parameter description", "description"], 5)
    i_ref = col_idx(["reference"], 6)
    i_363 = col_idx(["363"], 7)
    i_376 = col_idx(["376"], 8)
    i_j3u = col_idx(["j3u"], 9)
    i_hw = col_idx(["column used by ipc hw", "ipc hw"], 10)

    rows_out: list[SheetParameterRow] = []
    for r_idx, row in enumerate(selected_sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not row or len(row) <= i_name:
            continue
        p_name = str(row[i_name] or "").strip()
        if not p_name or p_name.lower() in ("none", "-"):
            continue
        ftv = str(row[i_ftv] if len(row) > i_ftv and row[i_ftv] is not None else "").strip()
        ref = str(row[i_ref] if len(row) > i_ref and row[i_ref] is not None else "").strip()
        rows_out.append(
            SheetParameterRow(
                sheet_name=selected_sheet.title,
                row_number=r_idx,
                name=p_name,
                normalized_name=normalize_param_name(p_name),
                first_trial_value=ftv,
                range_value=str(row[i_range] if len(row) > i_range and row[i_range] is not None else "").strip(),
                resolution=str(row[i_res] if len(row) > i_res and row[i_res] is not None else "").strip(),
                unit=str(row[i_unit] if len(row) > i_unit and row[i_unit] is not None else "").strip(),
                description=str(row[i_desc] if len(row) > i_desc and row[i_desc] is not None else "").strip(),
                reference=ref,
                extracted_vf_ids=extract_vf_ids(ref),
                p363=str(row[i_363] if len(row) > i_363 and row[i_363] is not None else "").strip(),
                p376=str(row[i_376] if len(row) > i_376 and row[i_376] is not None else "").strip(),
                pj3u=str(row[i_j3u] if len(row) > i_j3u and row[i_j3u] is not None else "").strip(),
                hw_col=str(row[i_hw] if len(row) > i_hw and row[i_hw] is not None else "").strip(),
                source_file=path.name,
            )
        )

    wb.close()
    return rows_out


def discover_and_parse_project(root_folder: str | Path, prefer_format: str = "docx") -> dict:
    """
    Dado o diretório raiz de um projeto, localiza automaticamente:
    - Arquivos de VF (.docx ou .pdf)
    - Planilha Indications_*.xlsx
    - Planilha *Parameters_Table*.xlsx
    """
    root = Path(root_folder)
    if not root.exists():
        raise FileNotFoundError(f"Diretório não encontrado: {root}")

    indications_files = []
    parameters_files = []
    docx_files = []
    pdf_files = []

    for f in root.rglob("*"):
        if not f.is_file() or f.name.startswith("~$"):
            continue
        low = f.name.lower()
        if low.endswith(".xlsx"):
            if "indication" in low:
                indications_files.append(f)
            elif "parameter" in low:
                parameters_files.append(f)
        elif low.endswith(".docx"):
            docx_files.append(f)
        elif low.endswith(".pdf") and "relatorio" not in low and "audit" not in low:
            pdf_files.append(f)

    if not indications_files:
        raise FileNotFoundError(f"Nenhuma planilha 'Indications*.xlsx' encontrada em {root}")
    if not parameters_files:
        raise FileNotFoundError(f"Nenhuma planilha '*Parameters_Table*.xlsx' encontrada em {root}")

    indications_path = sorted(indications_files)[0]
    parameters_path = sorted(parameters_files)[0]

    candidate_vf_files = docx_files if (prefer_format == "docx" and docx_files) else (pdf_files or docx_files)
    if not candidate_vf_files:
        raise FileNotFoundError(f"Nenhum arquivo de VF (.docx ou .pdf) encontrado em {root}")

    individual_candidates = [f for f in candidate_vf_files if re.match(r"^VF\d+", f.name, re.IGNORECASE)]
    files_to_parse = sorted(individual_candidates) if individual_candidates else sorted(candidate_vf_files)

    vfs_by_id: dict[str, ParsedVFDocument] = {}
    parsed_file_names = []
    total_pdf_pages = 0
    for vf_file in files_to_parse:
        parsed_file_names.append(vf_file.name)
        if vf_file.suffix.lower() == ".docx":
            parsed_vfs = parse_vf_docx(vf_file)
        else:
            with fitz.open(str(vf_file)) as pdoc:
                total_pdf_pages += len(pdoc)
            parsed_vfs = parse_vf_pdf(vf_file)
        for p_vf in parsed_vfs:
            vfs_by_id[p_vf.vf_id] = p_vf

    indications_rows = parse_indications_excel(indications_path)
    parameters_rows = parse_parameters_excel(parameters_path)

    return {
        "project_root": str(root),
        "vf_files": parsed_file_names,
        "total_pdf_pages": total_pdf_pages,
        "indications_file": indications_path.name,
        "indications_path": str(indications_path),
        "parameters_file": parameters_path.name,
        "parameters_path": str(parameters_path),
        "vfs": list(vfs_by_id.values()),
        "indications_rows": indications_rows,
        "parameters_rows": parameters_rows,
    }
