"""
Motor de Auditoria Bidirecional de Consistência (Prevenção de RARs - Stellantis IPC):
- Regra 1 (Bidirecional): VF (Indication Management) <-> Planilha Section C (Indications.xlsx)
- Regra 2 (Bidirecional): VF (Configuration Parameters Table [IPC + SUPPLIER EOL]) <-> Planilha Section B (Parameters_Table.xlsx)
"""

from datetime import datetime
from typing import Any
from app.models import (
    ParsedVFDocument,
    SheetIndicationRow,
    SheetParameterRow,
    AuditFinding,
    AuditReport,
    VFIndicationItem,
    VFParameterItem,
)


def normalize_numeric_str(val: str) -> str:
    """Normaliza valores numéricos com vírgula ou ponto decimal (ex: '2,5' == '2.5', '4,0' == '4') para comparação justa."""
    if not val:
        return ""
    s = str(val).strip().replace(",", ".")
    if s in ("-", "N/A", "NONE", "NULL"):
        return ""
    try:
        num = float(s)
        return f"{num:g}"
    except ValueError:
        return s.upper()


def run_bidirectional_audit(
    vfs: list[ParsedVFDocument],
    indications_rows: list[SheetIndicationRow],
    parameters_rows: list[SheetParameterRow],
    project_name: str = "J3U - 5870 (IPC - Instrument Panel Cluster 10 CAN)",
    files_info: dict[str, Any] | None = None,
) -> AuditReport:
    # Indexar indicações da planilha por normalized_name
    sheet_ind_by_norm: dict[str, list[SheetIndicationRow]] = {}
    for row in indications_rows:
        sheet_ind_by_norm.setdefault(row.normalized_name, []).append(row)

    # Indexar indicações das VFs por normalized_name
    vf_ind_by_norm: dict[str, list[VFIndicationItem]] = {}
    all_vf_indications: list[VFIndicationItem] = []
    for vf in vfs:
        for ind in vf.indications:
            all_vf_indications.append(ind)
            vf_ind_by_norm.setdefault(ind.normalized_name, []).append(ind)

    rule1_findings: list[AuditFinding] = []

    # =========================================================================
    # REGRA 1 - DIREÇÃO A: VF -> Planilha Indications
    # =========================================================================
    for ind in all_vf_indications:
        matching_sheet_rows = sheet_ind_by_norm.get(ind.normalized_name, [])
        if not matching_sheet_rows:
            rule1_findings.append(
                AuditFinding(
                    rule="RULE_1_INDICATIONS",
                    direction="VF_TO_SHEET",
                    status="MISSING_IN_SHEET",
                    severity="HIGH_RAR_RISK",
                    item_name=ind.raw_name,
                    normalized_name=ind.normalized_name,
                    vf_id=ind.vf_id,
                    sheet_vf_ref="AUSENTE",
                    details=(
                        f"A indicação '{ind.raw_name}' é requisitada na seção Indication Management da {ind.vf_id} "
                        f"({ind.vf_full_code}), mas NÃO existe na planilha Indications."
                    ),
                    recommendation=(
                        f"Incluir '{ind.normalized_name}' na planilha Indications vinculada à '{ind.vf_full_code}' "
                        f"ou revisar a {ind.vf_id} com o time de Sistema antes da entrega ao fornecedor."
                    ),
                    extra={
                        "vf_title": ind.vf_title,
                        "condition": ind.condition,
                        "popup": ind.popup_message,
                        "telltale": ind.telltale_behavior,
                        "source_file": ind.source_file,
                    },
                )
            )
        else:
            # Verificar se alguma linha correspondente aponta para a mesma VF (ex: VF143 em VF143_V1_PJ3U)
            s_row = matching_sheet_rows[0]
            sheet_vfs_str = s_row.vfs_raw or "(Vazio)"
            if ind.vf_id in s_row.extracted_vf_ids:
                rule1_findings.append(
                    AuditFinding(
                        rule="RULE_1_INDICATIONS",
                        direction="VF_TO_SHEET",
                        status="OK",
                        severity="OK",
                        item_name=ind.raw_name,
                        normalized_name=ind.normalized_name,
                        vf_id=ind.vf_id,
                        sheet_vf_ref=sheet_vfs_str,
                        details=(
                            f"Indicação '{ind.raw_name}' validada com sucesso na planilha Indications "
                            f"(Linha {s_row.row_number}, Nome na planilha: '{s_row.raw_name}', VF: '{sheet_vfs_str}')."
                        ),
                        recommendation="Nenhuma ação necessária.",
                        extra={
                            "sheet_row": s_row.row_number,
                            "sheet_name": s_row.sheet_name,
                            "hard_tt": s_row.hard_tt_code,
                            "display_msg": s_row.display_messages,
                        },
                    )
                )
            else:
                rule1_findings.append(
                    AuditFinding(
                        rule="RULE_1_INDICATIONS",
                        direction="VF_TO_SHEET",
                        status="VF_MISMATCH",
                        severity="MEDIUM_RAR_RISK",
                        item_name=ind.raw_name,
                        normalized_name=ind.normalized_name,
                        vf_id=ind.vf_id,
                        sheet_vf_ref=sheet_vfs_str,
                        details=(
                            f"A indicação '{ind.raw_name}' é chamada na {ind.vf_id} ({ind.vf_full_code}), "
                            f"mas na planilha Indications (Linha {s_row.row_number}) está vinculada à VF '{sheet_vfs_str}'."
                        ),
                        recommendation=(
                            f"Atualizar a coluna 'VFs' na linha {s_row.row_number} da planilha Indications de "
                            f"'{sheet_vfs_str}' para '{ind.vf_full_code}'."
                        ),
                        extra={
                            "sheet_row": s_row.row_number,
                            "sheet_name": s_row.sheet_name,
                        },
                    )
                )

    # =========================================================================
    # REGRA 1 - DIREÇÃO B: Planilha Indications -> VF (Detecção de Órfãos na Planilha)
    # =========================================================================
    for s_row in indications_rows:
        matching_vf_items = vf_ind_by_norm.get(s_row.normalized_name, [])
        sheet_vfs_str = s_row.vfs_raw or "(Sem VF preenchida)"
        if not matching_vf_items:
            rule1_findings.append(
                AuditFinding(
                    rule="RULE_1_INDICATIONS",
                    direction="SHEET_TO_VF",
                    status="ORPHAN_IN_SHEET",
                    severity="HIGH_RAR_RISK",
                    item_name=s_row.raw_name,
                    normalized_name=s_row.normalized_name,
                    vf_id="NENHUMA",
                    sheet_vf_ref=sheet_vfs_str,
                    details=(
                        f"Indicação ÓRFÃ na planilha Indications (Linha {s_row.row_number}, Coluna VFs: '{sheet_vfs_str}'): "
                        f"'{s_row.raw_name}' consta na planilha do componente IPC, mas NÃO é chamada por nenhuma VF do pacote."
                    ),
                    recommendation=(
                        f"Remover a linha {s_row.row_number} ('{s_row.raw_name}') caso seja herança (carryover) indevida "
                        f"de outro projeto, ou solicitar a inclusão da respectiva VF no pacote IPC."
                    ),
                    extra={
                        "sheet_row": s_row.row_number,
                        "sheet_name": s_row.sheet_name,
                        "hard_tt": s_row.hard_tt_code,
                        "display_msg": s_row.display_messages,
                    },
                )
            )
        else:
            calling_vfs = [item.vf_id for item in matching_vf_items]
            if any(vf_id in s_row.extracted_vf_ids for vf_id in calling_vfs):
                rule1_findings.append(
                    AuditFinding(
                        rule="RULE_1_INDICATIONS",
                        direction="SHEET_TO_VF",
                        status="OK",
                        severity="OK",
                        item_name=s_row.raw_name,
                        normalized_name=s_row.normalized_name,
                        vf_id=", ".join(calling_vfs),
                        sheet_vf_ref=sheet_vfs_str,
                        details=(
                            f"Linha {s_row.row_number} da planilha Indications ('{s_row.raw_name}') possui correspondência "
                            f"exata com a {', '.join(calling_vfs)}."
                        ),
                        recommendation="Nenhuma ação necessária.",
                        extra={"sheet_row": s_row.row_number, "sheet_name": s_row.sheet_name},
                    )
                )
            else:
                rule1_findings.append(
                    AuditFinding(
                        rule="RULE_1_INDICATIONS",
                        direction="SHEET_TO_VF",
                        status="VF_MISMATCH",
                        severity="MEDIUM_RAR_RISK",
                        item_name=s_row.raw_name,
                        normalized_name=s_row.normalized_name,
                        vf_id=", ".join(calling_vfs),
                        sheet_vf_ref=sheet_vfs_str,
                        details=(
                            f"Linha {s_row.row_number} da planilha Indications ('{s_row.raw_name}') referencia '{sheet_vfs_str}', "
                            f"mas quem chama esta indicação no pacote é {', '.join(calling_vfs)}."
                        ),
                        recommendation=f"Corrigir a coluna 'VFs' na linha {s_row.row_number} da planilha Indications para {', '.join(calling_vfs)}.",
                        extra={"sheet_row": s_row.row_number, "sheet_name": s_row.sheet_name},
                    )
                )

    # =========================================================================
    # REGRA 2 - SEPARAÇÃO E FILTRAGEM DE PARÂMETROS (IPC + SUPPLIER EOL)
    # =========================================================================
    in_scope_vf_params: list[VFParameterItem] = []
    filtered_out_params: list[dict[str, Any]] = []
    vf_param_by_norm: dict[str, list[VFParameterItem]] = {}

    for vf in vfs:
        for par in vf.parameters:
            if par.is_ipc_supplier_eol:
                in_scope_vf_params.append(par)
                vf_param_by_norm.setdefault(par.normalized_name, []).append(par)
            else:
                filtered_out_params.append(
                    {
                        "vf_id": par.vf_id,
                        "vf_full_code": par.vf_full_code,
                        "param_id": par.param_id,
                        "name": par.name,
                        "first_trial_value": par.first_trial_value,
                        "unit": par.unit,
                        "component": par.component,
                        "implementation": par.implementation,
                        "filter_reason": par.filter_reason,
                        "description": par.description,
                    }
                )

    sheet_param_by_norm: dict[str, list[SheetParameterRow]] = {}
    for p_row in parameters_rows:
        sheet_param_by_norm.setdefault(p_row.normalized_name, []).append(p_row)

    rule2_findings: list[AuditFinding] = []

    # =========================================================================
    # REGRA 2 - DIREÇÃO A: VF (IPC + SUPPLIER EOL) -> Parameters Table
    # =========================================================================
    for par in in_scope_vf_params:
        matching_p_rows = sheet_param_by_norm.get(par.normalized_name, [])
        if not matching_p_rows:
            rule2_findings.append(
                AuditFinding(
                    rule="RULE_2_PARAMETERS",
                    direction="VF_TO_SHEET",
                    status="MISSING_IN_SHEET",
                    severity="HIGH_RAR_RISK",
                    item_name=par.name,
                    normalized_name=par.normalized_name,
                    vf_id=par.vf_id,
                    sheet_vf_ref="AUSENTE",
                    details=(
                        f"O parâmetro '{par.name}' (Component='{par.component}', Implementation='{par.implementation}', "
                        f"First Trial='{par.first_trial_value} {par.unit}') é exigido na {par.vf_id}, mas NÃO existe na planilha Parameters Table."
                    ),
                    recommendation=(
                        f"Adicionar o parâmetro '{par.name}' na planilha IPC_Parameters_Table com First Trial Value = "
                        f"'{par.first_trial_value}', Range = '{par.range_value}', Resolution = '{par.resolution}', Unit = '{par.unit}'."
                    ),
                    extra={
                        "vf_first_trial": par.first_trial_value,
                        "sheet_first_trial": "AUSENTE",
                        "range_value": par.range_value,
                        "resolution": par.resolution,
                        "unit": par.unit,
                        "component": par.component,
                        "implementation": par.implementation,
                        "description": par.description,
                    },
                )
            )
        else:
            p_row = matching_p_rows[0]
            vf_ftv_norm = normalize_numeric_str(par.first_trial_value)
            sheet_ftv_norm = normalize_numeric_str(p_row.first_trial_value)

            if not sheet_ftv_norm:
                rule2_findings.append(
                    AuditFinding(
                        rule="RULE_2_PARAMETERS",
                        direction="VF_TO_SHEET",
                        status="VALUE_MISMATCH_OR_EMPTY",
                        severity="MEDIUM_RAR_RISK",
                        item_name=par.name,
                        normalized_name=par.normalized_name,
                        vf_id=par.vf_id,
                        sheet_vf_ref=p_row.reference or "(Sem ref)",
                        details=(
                            f"O parâmetro '{par.name}' ({par.vf_id}) existe na Parameters Table (Linha {p_row.row_number}), "
                            f"porém a coluna 'First Trial Value' está EM BRANCO na planilha (na VF o valor é '{par.first_trial_value} {par.unit}')."
                        ),
                        recommendation=(
                            f"Preencher 'First Trial Value' = '{par.first_trial_value}' na linha {p_row.row_number} "
                            f"da planilha IPC_Parameters_Table."
                        ),
                        extra={
                            "sheet_row": p_row.row_number,
                            "vf_first_trial": par.first_trial_value,
                            "sheet_first_trial": "(Vazio)",
                            "unit": par.unit,
                            "component": par.component,
                            "implementation": par.implementation,
                        },
                    )
                )
            elif vf_ftv_norm != sheet_ftv_norm:
                rule2_findings.append(
                    AuditFinding(
                        rule="RULE_2_PARAMETERS",
                        direction="VF_TO_SHEET",
                        status="VALUE_MISMATCH_OR_EMPTY",
                        severity="MEDIUM_RAR_RISK",
                        item_name=par.name,
                        normalized_name=par.normalized_name,
                        vf_id=par.vf_id,
                        sheet_vf_ref=p_row.reference or "(Sem ref)",
                        details=(
                            f"Divergência de 'First Trial Value' para o parâmetro '{par.name}' ({par.vf_id}): "
                            f"na VF consta '{par.first_trial_value}', mas na Parameters Table (Linha {p_row.row_number}) consta '{p_row.first_trial_value}'."
                        ),
                        recommendation=(
                            f"Alinhar o valor de 'First Trial Value' do parâmetro '{par.name}' entre a {par.vf_id} "
                            f"('{par.first_trial_value}') e a linha {p_row.row_number} da Parameters Table ('{p_row.first_trial_value}')."
                        ),
                        extra={
                            "sheet_row": p_row.row_number,
                            "vf_first_trial": par.first_trial_value,
                            "sheet_first_trial": p_row.first_trial_value,
                            "unit": par.unit,
                            "component": par.component,
                            "implementation": par.implementation,
                        },
                    )
                )
            else:
                rule2_findings.append(
                    AuditFinding(
                        rule="RULE_2_PARAMETERS",
                        direction="VF_TO_SHEET",
                        status="OK",
                        severity="OK",
                        item_name=par.name,
                        normalized_name=par.normalized_name,
                        vf_id=par.vf_id,
                        sheet_vf_ref=p_row.reference or par.vf_id,
                        details=(
                            f"Parâmetro '{par.name}' ({par.vf_id}, {par.component} / {par.implementation}) validado na Parameters Table "
                            f"(Linha {p_row.row_number}, First Trial Value = '{p_row.first_trial_value} {p_row.unit}')."
                        ),
                        recommendation="Nenhuma ação necessária.",
                        extra={
                            "sheet_row": p_row.row_number,
                            "vf_first_trial": par.first_trial_value,
                            "sheet_first_trial": p_row.first_trial_value,
                            "unit": par.unit,
                            "component": par.component,
                            "implementation": par.implementation,
                        },
                    )
                )

    # =========================================================================
    # REGRA 2 - DIREÇÃO B: Parameters Table -> VF (Detecção de Parâmetros Órfãos)
    # =========================================================================
    for p_row in parameters_rows:
        matching_vf_pars = vf_param_by_norm.get(p_row.normalized_name, [])
        if not matching_vf_pars:
            rule2_findings.append(
                AuditFinding(
                    rule="RULE_2_PARAMETERS",
                    direction="SHEET_TO_VF",
                    status="ORPHAN_IN_SHEET",
                    severity="HIGH_RAR_RISK",
                    item_name=p_row.name,
                    normalized_name=p_row.normalized_name,
                    vf_id="NENHUMA",
                    sheet_vf_ref=p_row.reference or "(Sem referência)",
                    details=(
                        f"Parâmetro ÓRFÃO na planilha Parameters Table (Linha {p_row.row_number}, Ref: '{p_row.reference}'): "
                        f"'{p_row.name}' está listado na especificação do IPC, mas NÃO é utilizado por nenhuma VF do pacote."
                    ),
                    recommendation=(
                        f"Remover o parâmetro órfão '{p_row.name}' (linha {p_row.row_number}) da planilha IPC_Parameters_Table "
                        f"ou verificar se falta incluir a VF correspondente no pacote."
                    ),
                    extra={
                        "sheet_row": p_row.row_number,
                        "vf_first_trial": "N/A",
                        "sheet_first_trial": p_row.first_trial_value,
                        "unit": p_row.unit,
                        "description": p_row.description,
                    },
                )
            )
        else:
            calling_vfs = [vp.vf_id for vp in matching_vf_pars]
            sheet_ftv_norm = normalize_numeric_str(p_row.first_trial_value)
            vf_ftv_norm = normalize_numeric_str(matching_vf_pars[0].first_trial_value)

            if not sheet_ftv_norm or sheet_ftv_norm != vf_ftv_norm:
                rule2_findings.append(
                    AuditFinding(
                        rule="RULE_2_PARAMETERS",
                        direction="SHEET_TO_VF",
                        status="VALUE_MISMATCH_OR_EMPTY",
                        severity="MEDIUM_RAR_RISK",
                        item_name=p_row.name,
                        normalized_name=p_row.normalized_name,
                        vf_id=", ".join(calling_vfs),
                        sheet_vf_ref=p_row.reference,
                        details=(
                            f"Linha {p_row.row_number} da Parameters Table ('{p_row.name}') é utilizada pela {', '.join(calling_vfs)}, "
                            f"mas possui 'First Trial Value' ausente ou divergente (Planilha: '{p_row.first_trial_value or '(Vazio)'}' vs VF: '{matching_vf_pars[0].first_trial_value}')."
                        ),
                        recommendation=f"Atualizar 'First Trial Value' na linha {p_row.row_number} para '{matching_vf_pars[0].first_trial_value}'.",
                        extra={
                            "sheet_row": p_row.row_number,
                            "vf_first_trial": matching_vf_pars[0].first_trial_value,
                            "sheet_first_trial": p_row.first_trial_value or "(Vazio)",
                            "unit": p_row.unit,
                        },
                    )
                )
            else:
                rule2_findings.append(
                    AuditFinding(
                        rule="RULE_2_PARAMETERS",
                        direction="SHEET_TO_VF",
                        status="OK",
                        severity="OK",
                        item_name=p_row.name,
                        normalized_name=p_row.normalized_name,
                        vf_id=", ".join(calling_vfs),
                        sheet_vf_ref=p_row.reference,
                        details=(
                            f"Linha {p_row.row_number} da Parameters Table ('{p_row.name}') é referenciada corretamente "
                            f"pela {', '.join(calling_vfs)}."
                        ),
                        recommendation="Nenhuma ação necessária.",
                        extra={
                            "sheet_row": p_row.row_number,
                            "vf_first_trial": matching_vf_pars[0].first_trial_value,
                            "sheet_first_trial": p_row.first_trial_value,
                            "unit": p_row.unit,
                        },
                    )
                )

    # Resumo por VF para o explorador de VFs
    vfs_overview = []
    for vf in vfs:
        vf_r1_issues = [
            f for f in rule1_findings if f.vf_id == vf.vf_id and f.direction == "VF_TO_SHEET" and f.status != "OK"
        ]
        vf_r2_issues = [
            f for f in rule2_findings if f.vf_id == vf.vf_id and f.direction == "VF_TO_SHEET" and f.status != "OK"
        ]
        in_scope_cnt = sum(1 for p in vf.parameters if p.is_ipc_supplier_eol)
        out_scope_cnt = sum(1 for p in vf.parameters if not p.is_ipc_supplier_eol)
        vfs_overview.append(
            {
                "vf_id": vf.vf_id,
                "vf_full_code": vf.vf_full_code,
                "title": vf.title,
                "source_file": vf.source_file,
                "indications_count": len(vf.indications),
                "parameters_total": len(vf.parameters),
                "parameters_ipc_eol": in_scope_cnt,
                "parameters_filtered_out": out_scope_cnt,
                "indications_issues": len(vf_r1_issues),
                "parameters_issues": len(vf_r2_issues),
                "status": "OK" if (len(vf_r1_issues) == 0 and len(vf_r2_issues) == 0) else "ATTENTION",
                "indications": [
                    {
                        "raw_name": i.raw_name,
                        "normalized_name": i.normalized_name,
                        "condition": i.condition,
                        "popup": i.popup_message,
                        "telltale": i.telltale_behavior,
                    }
                    for i in vf.indications
                ],
                "parameters": [
                    {
                        "param_id": p.param_id,
                        "name": p.name,
                        "first_trial_value": p.first_trial_value,
                        "range_value": p.range_value,
                        "resolution": p.resolution,
                        "unit": p.unit,
                        "component": p.component,
                        "implementation": p.implementation,
                        "is_ipc_supplier_eol": p.is_ipc_supplier_eol,
                        "filter_reason": p.filter_reason,
                    }
                    for p in vf.parameters
                ],
            }
        )

    # Calcular KPIs executivos
    r1_anomalies = [f for f in rule1_findings if f.status != "OK"]
    r2_anomalies = [f for f in rule2_findings if f.status != "OK"]

    # Contabilizar inconsistências únicas para evitar dupla contagem do mesmo VF_MISMATCH ou VALUE_MISMATCH nas duas direções
    r1_missing_in_sheet = sum(1 for f in rule1_findings if f.direction == "VF_TO_SHEET" and f.status == "MISSING_IN_SHEET")
    r1_orphan_in_sheet = sum(1 for f in rule1_findings if f.direction == "SHEET_TO_VF" and f.status == "ORPHAN_IN_SHEET")
    r1_vf_mismatch = sum(1 for f in rule1_findings if f.direction == "VF_TO_SHEET" and f.status == "VF_MISMATCH")
    r1_unique_issues = r1_missing_in_sheet + r1_orphan_in_sheet + r1_vf_mismatch

    r2_missing_in_sheet = sum(1 for f in rule2_findings if f.direction == "VF_TO_SHEET" and f.status == "MISSING_IN_SHEET")
    r2_orphan_in_sheet = sum(1 for f in rule2_findings if f.direction == "SHEET_TO_VF" and f.status == "ORPHAN_IN_SHEET")
    r2_value_mismatch = sum(
        1 for f in rule2_findings if f.direction == "VF_TO_SHEET" and f.status == "VALUE_MISMATCH_OR_EMPTY"
    )
    r2_unique_issues = r2_missing_in_sheet + r2_orphan_in_sheet + r2_value_mismatch

    total_unique_rar_risks = r1_unique_issues + r2_unique_issues
    total_checks = len(all_vf_indications) + len(indications_rows) + len(in_scope_vf_params) + len(parameters_rows)
    ok_checks = sum(1 for f in rule1_findings if f.status == "OK") + sum(1 for f in rule2_findings if f.status == "OK")
    conformity_rate = round((ok_checks / total_checks) * 100, 1) if total_checks > 0 else 100.0

    summary_kpis = {
        "total_vfs_analyzed": len(vfs),
        "total_vf_indications": len(all_vf_indications),
        "total_sheet_indications": len(indications_rows),
        "total_vf_params_all": len(in_scope_vf_params) + len(filtered_out_params),
        "total_vf_params_ipc_eol": len(in_scope_vf_params),
        "total_vf_params_filtered_out": len(filtered_out_params),
        "total_sheet_params": len(parameters_rows),
        "rule1_unique_issues": r1_unique_issues,
        "rule1_missing_in_sheet": r1_missing_in_sheet,
        "rule1_orphan_in_sheet": r1_orphan_in_sheet,
        "rule1_vf_mismatch": r1_vf_mismatch,
        "rule2_unique_issues": r2_unique_issues,
        "rule2_missing_in_sheet": r2_missing_in_sheet,
        "rule2_orphan_in_sheet": r2_orphan_in_sheet,
        "rule2_value_mismatch": r2_value_mismatch,
        "total_unique_rar_risks": total_unique_rar_risks,
        "conformity_rate_pct": conformity_rate,
        "total_checks_performed": total_checks,
        "total_anomalies_both_directions": len(r1_anomalies) + len(r2_anomalies),
    }

    return AuditReport(
        project_name=project_name,
        timestamp=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        files_analyzed=files_info or {},
        summary_kpis=summary_kpis,
        rule1_findings=rule1_findings,
        rule2_findings=rule2_findings,
        filtered_out_parameters=filtered_out_params,
        vfs_overview=vfs_overview,
    )
