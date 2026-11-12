"""
Suíte de testes automatizados (pytest) para o Stellantis IPC Consistency Auditor.
Valida:
1. Normalização de nomenclatura (.V1 vs ' V1') e extração de códigos VF.
2. Filtragem estrita de parâmetros (Component = IPC & Implementation = SUPPLIER EOL vs PROXI / outras ECUs).
3. Leitura de VFs tanto em .DOCX quanto em .PDF (individuais e consolidadas).
4. Detecção exata de todos os 10 cenários de RAR nos dados fictícios (J3U - 5870).
5. Endpoints da API FastAPI e exportação de relatórios em Excel (.xlsx) e PDF (.pdf).
"""

from pathlib import Path
from fastapi.testclient import TestClient

from app.parsers import (
    normalize_indication_name,
    extract_vf_ids,
    evaluate_ipc_supplier_eol_filter,
    discover_and_parse_project,
    parse_vf_docx,
    parse_vf_pdf,
)
from app.auditor import run_bidirectional_audit
from app.main import app
from scripts.generate_mock_data import MOCK_ROOT, VF_DIR, generate_all_mock_files


client = TestClient(app)


def setup_module():
    if not MOCK_ROOT.exists():
        generate_all_mock_files()


def test_indication_normalization():
    assert normalize_indication_name("HAND BRAKE.V1") == "HAND BRAKE V1"
    assert normalize_indication_name("  air bag temporary calibration.v1 ") == "AIR BAG TEMPORARY CALIBRATION V1"
    assert normalize_indication_name("HAND BRAKE VEHICLE IN MOVEMENT V1") == "HAND BRAKE VEHICLE IN MOVEMENT V1"


def test_vf_id_extraction():
    assert extract_vf_ids("VF143_V1_PJ3U") == ["VF143"]
    assert extract_vf_ids("62 VF143 - Hand Brake Management (VF143_V1_PJ3U)") == ["VF143"]
    assert extract_vf_ids("") == []


def test_ipc_supplier_eol_filter():
    in_scope, _ = evaluate_ipc_supplier_eol_filter("IPC", "SUPPLIER EOL_2")
    assert in_scope is True

    in_scope_proxi, reason_proxi = evaluate_ipc_supplier_eol_filter("IPC", "PROXI")
    assert in_scope_proxi is False
    assert "PROXI" in reason_proxi

    in_scope_orc, reason_orc = evaluate_ipc_supplier_eol_filter("ORC", "SUPPLIER EOL")
    assert in_scope_orc is False
    assert "ORC" in reason_orc


def test_consolidated_docx_and_pdf_parsers_match():
    docx_vfs = parse_vf_docx(VF_DIR / "VFS_IPC_PJ3U_2A.docx")
    pdf_vfs = parse_vf_pdf(VF_DIR / "VFS_IPC_PJ3U_2A.pdf")

    assert len(docx_vfs) == 6
    assert len(pdf_vfs) == 6

    docx_ids = [v.vf_id for v in docx_vfs]
    pdf_ids = [v.vf_id for v in pdf_vfs]
    assert docx_ids == pdf_ids == ["VF139", "VF143", "VF145", "VF150", "VF118", "VF210"]

    for dv, pv in zip(docx_vfs, pdf_vfs):
        assert len(dv.indications) == len(pv.indications)
        assert len(dv.parameters) == len(pv.parameters)


def test_bidirectional_audit_detects_all_10_rar_scenarios():
    parsed = discover_and_parse_project(MOCK_ROOT, prefer_format="docx")
    report = run_bidirectional_audit(
        vfs=parsed["vfs"],
        indications_rows=parsed["indications_rows"],
        parameters_rows=parsed["parameters_rows"],
    )
    kpis = report.summary_kpis

    # 6 VFs analisadas
    assert kpis["total_vfs_analyzed"] == 6

    # Regra 1 (VF <-> Indications):
    # - 2 ausentes na planilha: PASSENGER AIR BAG OFF.V1 (VF139) e EPB AUTO HOLD ACTIVE.V1 (VF143)
    # - 2 órfãs na planilha: GLOW PLUG PRE-HEATING V1 (VF199) e HILL DESCENT CONTROL FAIL V1 (sem VF)
    # - 1 divergência de VF: WINDSHIELD WASHER FLUID LOW.V1 (VF150 no doc vs VF149 na planilha)
    assert kpis["rule1_missing_in_sheet"] == 2
    assert kpis["rule1_orphan_in_sheet"] == 2
    assert kpis["rule1_vf_mismatch"] == 1
    assert kpis["rule1_unique_issues"] == 5

    # Regra 2 (VF <-> Parameters Table [IPC + SUPPLIER EOL]):
    # - 4 parâmetros filtrados fora do escopo (CFG_PAB_PRESENT [PROXI], T_ORC_CRASH_LATCH [ORC], T_BCM_HB_DEBOUNCE [BCM], CFG_REAR_SBR_ENABLE [PROXI])
    # - 2 ausentes na Parameters Table: T_HB_CHIME_REPEAT (VF143) e D_SBR_DIST_TH (VF118)
    # - 2 órfãos na Parameters Table: T_OIL_PRESS_DELAY_LEGACY e V_OVERSPEED_GCC_TH
    # - 1 com First Trial Value em branco na planilha: T_TPMS_BLINK_TIME (VF210)
    assert kpis["total_vf_params_filtered_out"] == 4
    assert kpis["rule2_missing_in_sheet"] == 2
    assert kpis["rule2_orphan_in_sheet"] == 2
    assert kpis["rule2_value_mismatch"] == 1
    assert kpis["rule2_unique_issues"] == 5

    # Total de 10 riscos únicos de RAR detectados
    assert kpis["total_unique_rar_risks"] == 10


def test_fastapi_endpoints_and_exports():
    # Dashboard inicial
    res_home = client.get("/")
    assert res_home.status_code == 200
    assert "Auditor de Especificações IPC" in res_home.text

    # Auditoria Mock DOCX
    res_docx = client.post("/api/audit/mock?prefer_format=docx")
    assert res_docx.status_code == 200
    data_docx = res_docx.json()
    assert data_docx["summary_kpis"]["total_unique_rar_risks"] == 10

    # Auditoria Mock PDF
    res_pdf = client.post("/api/audit/mock?prefer_format=pdf")
    assert res_pdf.status_code == 200
    data_pdf = res_pdf.json()
    assert data_pdf["summary_kpis"]["total_unique_rar_risks"] == 10

    # Auditoria via caminho de pasta
    res_folder = client.post(
        "/api/audit/folder",
        json={"folder_path": str(MOCK_ROOT), "prefer_format": "docx", "project_name": "Teste Pasta J3U"},
    )
    assert res_folder.status_code == 200
    assert res_folder.json()["summary_kpis"]["total_unique_rar_risks"] == 10

    # Exportação Excel e PDF
    res_exp_xlsx = client.get("/api/export/excel")
    assert res_exp_xlsx.status_code == 200
    assert len(res_exp_xlsx.content) > 4000

    res_exp_pdf = client.get("/api/export/pdf")
    assert res_exp_pdf.status_code == 200
    assert res_exp_pdf.content.startswith(b"%PDF")


def test_prose_doors_style_indication_extraction(tmp_path: Path):
    """Garante que o parser entende indicações escritas em texto corrido/bullets (padrão DOORS) mesmo sem tabela."""
    import docx

    doc = docx.Document()
    doc.add_heading("62 VF143 - Hand Brake Management (VF143_V1_PJ3U)", level=1)
    doc.add_heading("62.12 Indication Management", level=2)
    doc.add_paragraph(
        "When the parking brake switch is engaged and vehicle is stationary, the IPC shall activate the "
        "HAND BRAKE.V1 indication on the cluster."
    )
    doc.add_paragraph(
        "If vehicle speed exceeds V_TH_HB while engaged, the IPC shall trigger HAND BRAKE VEHICLE IN MOVEMENT.V1."
    )
    f_path = tmp_path / "VF143_Prose_DOORS.docx"
    doc.save(str(f_path))

    parsed = parse_vf_docx(f_path)
    assert len(parsed) == 1
    norms = [i.normalized_name for i in parsed[0].indications]
    assert "HAND BRAKE V1" in norms
    assert "HAND BRAKE VEHICLE IN MOVEMENT V1" in norms


def test_large_scale_1215_pages_vf_list_pdf_and_docx():
    """
    Valida o pacote de grande escala (1.215 páginas / 45 VFs completas com tabelas distratoras CAN/DTC
    e milhares de linhas de requisitos DOORS), garantindo precisão cirúrgica e alta performance.
    """
    import time
    from scripts.generate_large_scale_data import LARGE_ROOT, generate_large_scale_dataset

    if not LARGE_ROOT.exists():
        generate_large_scale_dataset()

    t0 = time.perf_counter()
    parsed_pdf = discover_and_parse_project(LARGE_ROOT, prefer_format="pdf")
    elapsed_pdf = time.perf_counter() - t0

    # Garantir que o PDF possui mais de 1.000 páginas (1.215 páginas) e foi lido rapidamente
    assert parsed_pdf["total_pdf_pages"] == 1215
    assert elapsed_pdf < 8.0

    report_pdf = run_bidirectional_audit(
        vfs=parsed_pdf["vfs"],
        indications_rows=parsed_pdf["indications_rows"],
        parameters_rows=parsed_pdf["parameters_rows"],
    )
    k = report_pdf.summary_kpis

    # 45 VFs analisadas, 135 indicações nas VFs (45 * 3), 80 parâmetros IPC+EOL e 80 filtrados (40 PROXI + 40 BCM)
    assert k["total_vfs_analyzed"] == 45
    assert k["total_vf_indications"] == 135
    assert k["total_vf_params_ipc_eol"] == 80
    assert k["total_vf_params_filtered_out"] == 80

    # Regra 1: 3 ausentes na planilha + 3 órfãs na planilha + 2 VF divergentes = 8 inconsistências
    assert k["rule1_missing_in_sheet"] == 3
    assert k["rule1_orphan_in_sheet"] == 3
    assert k["rule1_vf_mismatch"] == 2
    assert k["rule1_unique_issues"] == 8

    # Regra 2: 3 ausentes na planilha + 3 órfãos na planilha + 2 First Trial em branco = 8 inconsistências
    assert k["rule2_missing_in_sheet"] == 3
    assert k["rule2_orphan_in_sheet"] == 3
    assert k["rule2_value_mismatch"] == 2
    assert k["rule2_unique_issues"] == 8

    # Total de 16 riscos únicos de RAR em 430 checagens cruzadas
    assert k["total_unique_rar_risks"] == 16
    assert k["total_checks_performed"] == 430

    # Validar também via endpoint FastAPI /api/audit/large (tanto PDF quanto DOCX)
    res_large_pdf = client.post("/api/audit/large?prefer_format=pdf")
    assert res_large_pdf.status_code == 200
    assert res_large_pdf.json()["summary_kpis"]["total_unique_rar_risks"] == 16

    res_large_docx = client.post("/api/audit/large?prefer_format=docx")
    assert res_large_docx.status_code == 200
    assert res_large_docx.json()["summary_kpis"]["total_unique_rar_risks"] == 16


