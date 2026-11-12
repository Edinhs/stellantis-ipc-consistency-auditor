"""
Exportador de Relatórios de Auditoria Bidirecional Stellantis IPC:
- Exportação em Excel (.xlsx) multi-abas com formatação condicional de severidade RAR
- Exportação em PDF (.pdf) executivo pronto para encaminhamento aos times de Sistema e Componente
"""

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

from app.models import AuditReport, AuditFinding


STATUS_LABELS_PT = {
    "OK": "CONFORME (OK)",
    "MISSING_IN_SHEET": "AUSENTE NA PLANILHA (Risco RAR)",
    "ORPHAN_IN_SHEET": "ÓRFÃO NA PLANILHA (Risco RAR)",
    "VF_MISMATCH": "DIVERGÊNCIA DE VF (Risco RAR)",
    "VALUE_MISMATCH_OR_EMPTY": "VALOR FIRST TRIAL AUSENTE/DIVERGENTE",
}

DIRECTION_LABELS_PT = {
    "VF_TO_SHEET": "VF -> Planilha",
    "SHEET_TO_VF": "Planilha -> VF",
}


def generate_excel_report_bytes(report: AuditReport) -> bytes:
    wb = openpyxl.Workbook()

    # Estilos padrão Stellantis
    navy_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    blue_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    warn_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    err_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    header_font = Font(color="FFFFFF", bold=True, size=10)
    bold_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # 1. Aba Resumo Executivo (KPIs)
    ws_sum = wb.active
    ws_sum.title = "Resumo Executivo"
    ws_sum.append(["STELLANTIS - RELATÓRIO DE AUDITORIA DE CONSISTÊNCIA IPC (PREVENÇÃO DE RARs)", ""])
    ws_sum["A1"].font = Font(color="FFFFFF", bold=True, size=12)
    ws_sum["A1"].fill = navy_fill
    ws_sum["B1"].fill = navy_fill

    kpis = report.summary_kpis
    summary_rows = [
        ("Projeto Auditado", report.project_name),
        ("Data/Hora da Auditoria", report.timestamp),
        ("Taxa Global de Conformidade (%)", f"{kpis['conformity_rate_pct']}%"),
        ("Total de Riscos Únicos de RAR Detectados", kpis["total_unique_rar_risks"]),
        ("Total de VFs Analisadas no Pacote", kpis["total_vfs_analyzed"]),
        ("", ""),
        ("--- REGRA 1: VF <-> INDICATIONS (HMI Section C) ---", ""),
        ("Indicações Requisitadas nas VFs", kpis["total_vf_indications"]),
        ("Indicações Listadas na Planilha Indications", kpis["total_sheet_indications"]),
        ("Total de Inconsistências na Regra 1", kpis["rule1_unique_issues"]),
        ("  • Chamadas na VF, mas AUSENTES na planilha Indications", kpis["rule1_missing_in_sheet"]),
        ("  • ÓRFÃS na planilha Indications (sem VF correspondente)", kpis["rule1_orphan_in_sheet"]),
        ("  • Divergência de código de VF (VF Mismatch)", kpis["rule1_vf_mismatch"]),
        ("", ""),
        ("--- REGRA 2: VF <-> PARAMETERS TABLE (SW Section B - IPC + SUPPLIER EOL) ---", ""),
        ("Total de Parâmetros nas VFs (Todos os Componentes/Implementações)", kpis["total_vf_params_all"]),
        ("Parâmetros Auditáveis (Component = IPC & Implementation = SUPPLIER EOL)", kpis["total_vf_params_ipc_eol"]),
        ("Parâmetros Filtrados Fora do Escopo (PROXI / Outras ECUs)", kpis["total_vf_params_filtered_out"]),
        ("Parâmetros Listados na Planilha IPC_Parameters_Table", kpis["total_sheet_params"]),
        ("Total de Inconsistências na Regra 2", kpis["rule2_unique_issues"]),
        ("  • Exigidos na VF (IPC + SUPPLIER EOL), mas AUSENTES na Parameters Table", kpis["rule2_missing_in_sheet"]),
        ("  • ÓRFÃOS na Parameters Table (não usados por nenhuma VF)", kpis["rule2_orphan_in_sheet"]),
        ("  • First Trial Value em branco ou divergente", kpis["rule2_value_mismatch"]),
    ]

    for label, val in summary_rows:
        ws_sum.append([label, val])
        row_idx = ws_sum.max_row
        ws_sum.cell(row=row_idx, column=1).border = thin_border
        ws_sum.cell(row=row_idx, column=2).border = thin_border
        if label.startswith("---"):
            ws_sum.cell(row=row_idx, column=1).fill = blue_fill
            ws_sum.cell(row=row_idx, column=1).font = header_font
            ws_sum.cell(row=row_idx, column=2).fill = blue_fill
        elif label:
            ws_sum.cell(row=row_idx, column=1).font = bold_font

    ws_sum.column_dimensions["A"].width = 68
    ws_sum.column_dimensions["B"].width = 42

    # Função auxiliar para popular abas de Findings (Regra 1 e Regra 2)
    def populate_findings_sheet(ws, findings: list[AuditFinding], is_param: bool = False):
        headers = [
            "Direção da Checagem",
            "Status",
            "Severidade RAR",
            "Nome Original",
            "Nome Normalizado",
            "VF (Documento)",
            "Ref. VF (Planilha)",
        ]
        if is_param:
            headers.extend(["First Trial (VF)", "First Trial (Planilha)", "Unidade"])
        headers.extend(["Diagnóstico Detalhado", "Ação Recomendada (Prevenção RAR)"])

        ws.append(headers)
        for col_i in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=col_i)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Ordenar colocando inconsistências primeiro
        sorted_findings = sorted(findings, key=lambda x: (x.status == "OK", x.direction, x.vf_id, x.item_name))
        for f in sorted_findings:
            row_vals = [
                DIRECTION_LABELS_PT.get(f.direction, f.direction),
                STATUS_LABELS_PT.get(f.status, f.status),
                f.severity,
                f.item_name,
                f.normalized_name,
                f.vf_id,
                f.sheet_vf_ref,
            ]
            if is_param:
                row_vals.extend(
                    [
                        f.extra.get("vf_first_trial", "-"),
                        f.extra.get("sheet_first_trial", "-"),
                        f.extra.get("unit", "-"),
                    ]
                )
            row_vals.extend([f.details, f.recommendation])
            ws.append(row_vals)

            r_idx = ws.max_row
            fill_to_use = ok_fill if f.status == "OK" else (err_fill if f.severity == "HIGH_RAR_RISK" else warn_fill)
            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if c_idx in (2, 3):
                    cell.fill = fill_to_use
                    cell.font = bold_font

        for c_idx in range(1, len(headers) + 1):
            col_letter = openpyxl.utils.get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = 22 if c_idx < len(headers) - 1 else 52

    # 2. Aba Regra 1 (VF <-> Indications)
    ws_r1 = wb.create_sheet("Regra 1 - Indications")
    populate_findings_sheet(ws_r1, report.rule1_findings, is_param=False)

    # 3. Aba Regra 2 (VF <-> Parameters Table)
    ws_r2 = wb.create_sheet("Regra 2 - Parameters Table")
    populate_findings_sheet(ws_r2, report.rule2_findings, is_param=True)

    # 4. Aba Parâmetros Filtrados (PROXI & Outras ECUs)
    ws_filt = wb.create_sheet("Parâmetros Filtrados (PROXI)")
    filt_headers = [
        "VF",
        "Código VF Completo",
        "ID Parâmetro",
        "Parameter Name",
        "First Trial Value",
        "Unidade",
        "Component",
        "Implementation",
        "Motivo do Filtro (Fora do Escopo EOL IPC)",
    ]
    ws_filt.append(filt_headers)
    for col_i in range(1, len(filt_headers) + 1):
        c = ws_filt.cell(row=1, column=col_i)
        c.fill = blue_fill
        c.font = header_font
    for fp in report.filtered_out_parameters:
        ws_filt.append(
            [
                fp["vf_id"],
                fp["vf_full_code"],
                fp["param_id"],
                fp["name"],
                fp["first_trial_value"],
                fp["unit"],
                fp["component"],
                fp["implementation"],
                fp["filter_reason"],
            ]
        )
    for c_idx in range(1, len(filt_headers) + 1):
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws_filt.column_dimensions[col_letter].width = 24 if c_idx < 9 else 58

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_pdf_report_bytes(report: AuditReport) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Relatório de Auditoria IPC - {report.project_name}",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "StellantisTitle",
        parent=styles["Heading1"],
        fontSize=15,
        textColor=colors.HexColor("#1F3864"),
        spaceAfter=4,
    )
    sub_style = ParagraphStyle(
        "StellantisSub",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#444444"),
        spaceAfter=10,
    )
    sec_style = ParagraphStyle(
        "StellantisSec",
        parent=styles["Heading2"],
        fontSize=12,
        textColor=colors.HexColor("#2F5597"),
        spaceBefore=10,
        spaceAfter=6,
    )
    cell_style = ParagraphStyle(
        "CellText",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=9.5,
    )
    cell_hdr_style = ParagraphStyle(
        "CellHdrText",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        textColor=colors.white,
        fontName="Helvetica-Bold",
    )

    elements = []
    kpis = report.summary_kpis

    elements.append(Paragraph("STELLANTIS — RELATÓRIO DE AUDITORIA DE CONSISTÊNCIA IPC (PREVENÇÃO DE RARs)", title_style))
    elements.append(
        Paragraph(
            f"<b>Projeto:</b> {report.project_name} &nbsp;|&nbsp; <b>Data/Hora:</b> {report.timestamp} &nbsp;|&nbsp; "
            f"<b>Taxa de Conformidade:</b> {kpis['conformity_rate_pct']}% &nbsp;|&nbsp; "
            f"<b>Riscos Únicos de RAR Detectados:</b> {kpis['total_unique_rar_risks']}",
            sub_style,
        )
    )

    # Tabela Resumo KPI
    kpi_data = [
        [
            Paragraph("Métrica Auditada", cell_hdr_style),
            Paragraph("Valor", cell_hdr_style),
            Paragraph("Métrica Auditada", cell_hdr_style),
            Paragraph("Valor", cell_hdr_style),
        ],
        [
            Paragraph("Total de VFs Analisadas no Pacote", cell_style),
            Paragraph(str(kpis["total_vfs_analyzed"]), cell_style),
            Paragraph("Total de Riscos Únicos de RAR", cell_style),
            Paragraph(f"<b>{kpis['total_unique_rar_risks']}</b>", cell_style),
        ],
        [
            Paragraph("Regra 1: Indicações nas VFs vs Planilha", cell_style),
            Paragraph(f"{kpis['total_vf_indications']} vs {kpis['total_sheet_indications']}", cell_style),
            Paragraph("Regra 1: Ausentes / Órfãs / VF Divergente", cell_style),
            Paragraph(
                f"{kpis['rule1_missing_in_sheet']} ausentes | {kpis['rule1_orphan_in_sheet']} órfãs | {kpis['rule1_vf_mismatch']} VF divergente",
                cell_style,
            ),
        ],
        [
            Paragraph("Regra 2: Parâmetros IPC+SUPPLIER EOL vs Planilha", cell_style),
            Paragraph(
                f"{kpis['total_vf_params_ipc_eol']} EOL ({kpis['total_vf_params_filtered_out']} filtrados) vs {kpis['total_sheet_params']} na planilha",
                cell_style,
            ),
            Paragraph("Regra 2: Ausentes / Órfãos / Valor Divergente", cell_style),
            Paragraph(
                f"{kpis['rule2_missing_in_sheet']} ausentes | {kpis['rule2_orphan_in_sheet']} órfãos | {kpis['rule2_value_mismatch']} valor em branco",
                cell_style,
            ),
        ],
    ]
    kpi_table = Table(kpi_data, colWidths=[75 * mm, 55 * mm, 75 * mm, 65 * mm])
    kpi_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#F8FAFC"), colors.white]),
                ("PADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    elements.append(kpi_table)
    elements.append(Spacer(1, 6 * mm))

    # Seção 1: Inconsistências Regra 1 (VF <-> Indications)
    elements.append(Paragraph("1. Inconsistências e Checagens — Regra 1: VF ↔ Indications (HMI Section C)", sec_style))
    r1_rows = [
        [
            Paragraph("Direção", cell_hdr_style),
            Paragraph("Status", cell_hdr_style),
            Paragraph("Indicação", cell_hdr_style),
            Paragraph("VF Doc", cell_hdr_style),
            Paragraph("VF Planilha", cell_hdr_style),
            Paragraph("Diagnóstico", cell_hdr_style),
            Paragraph("Recomendação", cell_hdr_style),
        ]
    ]
    sorted_r1 = sorted(report.rule1_findings, key=lambda x: (x.status == "OK", x.direction, x.item_name))
    for f in sorted_r1:
        r1_rows.append(
            [
                Paragraph(DIRECTION_LABELS_PT.get(f.direction, f.direction), cell_style),
                Paragraph(STATUS_LABELS_PT.get(f.status, f.status), cell_style),
                Paragraph(f.item_name, cell_style),
                Paragraph(f.vf_id, cell_style),
                Paragraph(f.sheet_vf_ref, cell_style),
                Paragraph(f.details, cell_style),
                Paragraph(f.recommendation, cell_style),
            ]
        )

    r1_table = Table(r1_rows, colWidths=[26 * mm, 36 * mm, 42 * mm, 18 * mm, 26 * mm, 64 * mm, 60 * mm], repeatRows=1)
    r1_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("PADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elements.append(r1_table)
    elements.append(PageBreak())

    # Seção 2: Inconsistências Regra 2 (VF <-> Parameters Table)
    elements.append(
        Paragraph(
            "2. Inconsistências e Checagens — Regra 2: VF ↔ Parameters Table (Filtro: Component = IPC & Implementation = SUPPLIER EOL)",
            sec_style,
        )
    )
    r2_rows = [
        [
            Paragraph("Direção", cell_hdr_style),
            Paragraph("Status", cell_hdr_style),
            Paragraph("Parâmetro", cell_hdr_style),
            Paragraph("VF Doc", cell_hdr_style),
            Paragraph("Valor VF / Planilha", cell_hdr_style),
            Paragraph("Diagnóstico", cell_hdr_style),
            Paragraph("Recomendação", cell_hdr_style),
        ]
    ]
    sorted_r2 = sorted(report.rule2_findings, key=lambda x: (x.status == "OK", x.direction, x.item_name))
    for f in sorted_r2:
        val_comp = f"VF: {f.extra.get('vf_first_trial', '-')} | Plan: {f.extra.get('sheet_first_trial', '-')}"
        r2_rows.append(
            [
                Paragraph(DIRECTION_LABELS_PT.get(f.direction, f.direction), cell_style),
                Paragraph(STATUS_LABELS_PT.get(f.status, f.status), cell_style),
                Paragraph(f.item_name, cell_style),
                Paragraph(f.vf_id, cell_style),
                Paragraph(val_comp, cell_style),
                Paragraph(f.details, cell_style),
                Paragraph(f.recommendation, cell_style),
            ]
        )

    r2_table = Table(r2_rows, colWidths=[26 * mm, 36 * mm, 38 * mm, 18 * mm, 30 * mm, 64 * mm, 60 * mm], repeatRows=1)
    r2_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5597")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("PADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elements.append(r2_table)

    doc.build(elements)
    return buf.getvalue()
