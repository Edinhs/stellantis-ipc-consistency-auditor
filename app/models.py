from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class VFIndicationItem:
    vf_id: str
    vf_full_code: str
    vf_title: str
    raw_name: str
    normalized_name: str
    condition: str = ""
    popup_message: str = ""
    telltale_behavior: str = ""
    source_file: str = ""


@dataclass
class VFParameterItem:
    vf_id: str
    vf_full_code: str
    vf_title: str
    param_id: str
    name: str
    normalized_name: str
    first_trial_value: str
    range_value: str
    resolution: str
    unit: str
    component: str
    implementation: str
    xcp_availability: str = ""
    impact: str = ""
    description: str = ""
    is_ipc_supplier_eol: bool = False
    filter_reason: str = ""
    source_file: str = ""


@dataclass
class ParsedVFDocument:
    vf_id: str
    vf_full_code: str
    title: str
    source_file: str
    indications: list[VFIndicationItem] = field(default_factory=list)
    parameters: list[VFParameterItem] = field(default_factory=list)


@dataclass
class SheetIndicationRow:
    sheet_name: str
    row_number: int
    ipc_variant: str
    raw_name: str
    normalized_name: str
    vfs_raw: str
    extracted_vf_ids: list[str]
    hard_tt_code: str = ""
    soft_tt_code: str = ""
    driven: str = ""
    std_behav: str = ""
    display_messages: str = ""
    source_file: str = ""


@dataclass
class SheetParameterRow:
    sheet_name: str
    row_number: int
    name: str
    normalized_name: str
    first_trial_value: str
    range_value: str
    resolution: str
    unit: str
    description: str
    reference: str
    extracted_vf_ids: list[str]
    p363: str = ""
    p376: str = ""
    pj3u: str = ""
    hw_col: str = ""
    source_file: str = ""


@dataclass
class AuditFinding:
    rule: str  # "RULE_1_INDICATIONS" ou "RULE_2_PARAMETERS"
    direction: str  # "VF_TO_SHEET" ou "SHEET_TO_VF"
    status: str  # "OK", "MISSING_IN_SHEET", "ORPHAN_IN_SHEET", "VF_MISMATCH", "VALUE_MISMATCH_OR_EMPTY"
    severity: str  # "OK", "HIGH_RAR_RISK", "MEDIUM_RAR_RISK"
    item_name: str
    normalized_name: str
    vf_id: str
    sheet_vf_ref: str
    details: str
    recommendation: str
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class AuditReport:
    project_name: str
    timestamp: str
    files_analyzed: dict[str, Any]
    summary_kpis: dict[str, Any]
    rule1_findings: list[AuditFinding]
    rule2_findings: list[AuditFinding]
    filtered_out_parameters: list[dict[str, Any]]
    vfs_overview: list[dict[str, Any]]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
