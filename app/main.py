"""
Servidor Web FastAPI para o Stellantis IPC Consistency Auditor (Prevenção de RARs).
Oferece:
- Interface Web Interativa em Português (Dashboard Executivo e Técnico)
- Auditoria em 1 clique sobre a pasta de dados fictícios (J3U - 5870) em formato .docx ou .pdf
- Auditoria por caminho de pasta local/rede (Auto-Discovery na estrutura Stellantis)
- Auditoria via Upload direto de arquivos (.docx/.pdf + Indications.xlsx + Parameters_Table.xlsx)
- Exportação instantânea do relatório completo em Excel (.xlsx) e PDF (.pdf)
"""

from pathlib import Path
import tempfile
import shutil
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app.parsers import (
    discover_and_parse_project,
    parse_vf_docx,
    parse_vf_pdf,
    parse_indications_excel,
    parse_parameters_excel,
)
from app.auditor import run_bidirectional_audit
from app.exporter import generate_excel_report_bytes, generate_pdf_report_bytes
from app.models import AuditReport, ParsedVFDocument
from scripts.generate_mock_data import generate_all_mock_files, MOCK_ROOT
from scripts.generate_large_scale_data import generate_large_scale_dataset, LARGE_ROOT


BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = Path(__file__).resolve().parent / "static"
REPORTS_DIR = BASE_DIR / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(
    title="Stellantis IPC Consistency Auditor — Prevenção de RARs",
    description="Auditoria bidirecional automatizada entre VFs (Vehicle Functions) e especificações Extra-DOORS do IPC (Indications & Parameters Table).",
    version="1.0.0",
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Cache em memória do último relatório gerado para exportação rápida
_LAST_REPORT: AuditReport | None = None


class FolderAuditRequest(BaseModel):
    folder_path: str
    prefer_format: str = "docx"
    project_name: str = "Projeto IPC Auditado"


def _save_and_cache_report(report: AuditReport) -> AuditReport:
    global _LAST_REPORT
    _LAST_REPORT = report
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_bytes = generate_excel_report_bytes(report)
    pdf_bytes = generate_pdf_report_bytes(report)
    (REPORTS_DIR / "Relatorio_Auditoria_IPC_J3U.xlsx").write_bytes(xlsx_bytes)
    (REPORTS_DIR / "Relatorio_Auditoria_IPC_J3U.pdf").write_bytes(pdf_bytes)
    return report


@app.get("/", response_class=HTMLResponse)
def index():
    index_path = STATIC_DIR / "index.html"
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))


@app.post("/api/audit/mock")
def audit_mock_project(prefer_format: str = "docx"):
    """Garante que os arquivos fictícios (J3U - 5870) existam e executa a auditoria completa."""
    if not MOCK_ROOT.exists():
        generate_all_mock_files()

    parsed = discover_and_parse_project(MOCK_ROOT, prefer_format=prefer_format)
    report = run_bidirectional_audit(
        vfs=parsed["vfs"],
        indications_rows=parsed["indications_rows"],
        parameters_rows=parsed["parameters_rows"],
        project_name=f"J3U - 5870 | IPC - Instrument Panel Cluster 10 (CAN) [VFs .{prefer_format.upper()}]",
        files_info={
            "mode": f"Mock Dataset ({prefer_format.upper()})",
            "project_root": parsed["project_root"],
            "vf_files": parsed["vf_files"],
            "indications_file": parsed["indications_file"],
            "parameters_file": parsed["parameters_file"],
        },
    )
    _save_and_cache_report(report)
    return report.to_dict()


@app.post("/api/audit/large")
def audit_large_scale_project(prefer_format: str = "pdf"):
    """
    Executa a auditoria bidirecional no pacote de grande escala (1.215 páginas / 45 VFs completas),
    simulando o volume real de um documento VFS_IPC exportado do DOORS.
    """
    import time

    if not LARGE_ROOT.exists():
        generate_large_scale_dataset()

    t0 = time.perf_counter()
    parsed = discover_and_parse_project(LARGE_ROOT, prefer_format=prefer_format)
    report = run_bidirectional_audit(
        vfs=parsed["vfs"],
        indications_rows=parsed["indications_rows"],
        parameters_rows=parsed["parameters_rows"],
        project_name=(
            f"J3U - 5870 FULL PACKAGE (1.215 Páginas / 45 VFs) [.{prefer_format.upper()}]"
        ),
        files_info={
            "mode": f"Stress Test 1.215 Páginas ({prefer_format.upper()})",
            "project_root": parsed["project_root"],
            "vf_files": parsed["vf_files"],
            "total_pdf_pages": parsed.get("total_pdf_pages", 1215),
            "indications_file": parsed["indications_file"],
            "parameters_file": parsed["parameters_file"],
            "elapsed_seconds": round(time.perf_counter() - t0, 3),
        },
    )
    _save_and_cache_report(report)
    return report.to_dict()


@app.post("/api/audit/folder")
def audit_folder_project(req: FolderAuditRequest):
    """Varre automaticamente uma pasta local ou de rede no padrão Stellantis e executa a auditoria."""
    folder = Path(req.folder_path.strip().strip('"'))
    if not folder.exists():
        raise HTTPException(status_code=404, detail=f"Diretório não encontrado: {folder}")
    try:
        parsed = discover_and_parse_project(folder, prefer_format=req.prefer_format)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    report = run_bidirectional_audit(
        vfs=parsed["vfs"],
        indications_rows=parsed["indications_rows"],
        parameters_rows=parsed["parameters_rows"],
        project_name=req.project_name or folder.name,
        files_info={
            "mode": "Diretório Local / Rede",
            "project_root": parsed["project_root"],
            "vf_files": parsed["vf_files"],
            "indications_file": parsed["indications_file"],
            "parameters_file": parsed["parameters_file"],
        },
    )
    _save_and_cache_report(report)
    return report.to_dict()


@app.post("/api/audit/upload")
async def audit_uploaded_files(
    vf_files: list[UploadFile] = File(..., description="Um ou mais arquivos de VF (.docx ou .pdf)"),
    indications_file: UploadFile = File(..., description="Planilha Section C - Indications (.xlsx)"),
    parameters_file: UploadFile = File(..., description="Planilha Section B - Parameters Table (.xlsx)"),
    project_name: str = Form("Projeto IPC (Upload Manual)"),
):
    """Recebe os arquivos de VF (.docx/.pdf), Indications (.xlsx) e Parameters Table (.xlsx) via upload e executa a auditoria."""
    tmp_dir = Path(tempfile.mkdtemp(prefix="stellantis_ipc_audit_"))
    try:
        ind_path = tmp_dir / (indications_file.filename or "Indications.xlsx")
        ind_path.write_bytes(await indications_file.read())

        par_path = tmp_dir / (parameters_file.filename or "Parameters_Table.xlsx")
        par_path.write_bytes(await parameters_file.read())

        vfs_by_id: dict[str, ParsedVFDocument] = {}
        vf_names: list[str] = []

        for uf in vf_files:
            fname = uf.filename or "VF.docx"
            vf_names.append(fname)
            f_path = tmp_dir / fname
            f_path.write_bytes(await uf.read())
            if fname.lower().endswith(".docx"):
                parsed_vfs = parse_vf_docx(f_path)
            elif fname.lower().endswith(".pdf"):
                parsed_vfs = parse_vf_pdf(f_path)
            else:
                raise HTTPException(status_code=400, detail=f"Formato de VF não suportado: {fname}. Use .docx ou .pdf.")

            for p_vf in parsed_vfs:
                vfs_by_id[p_vf.vf_id] = p_vf

        ind_rows = parse_indications_excel(ind_path)
        par_rows = parse_parameters_excel(par_path)

        report = run_bidirectional_audit(
            vfs=list(vfs_by_id.values()),
            indications_rows=ind_rows,
            parameters_rows=par_rows,
            project_name=project_name,
            files_info={
                "mode": "Upload de Arquivos",
                "vf_files": vf_names,
                "indications_file": ind_path.name,
                "parameters_file": par_path.name,
            },
        )
        _save_and_cache_report(report)
        return report.to_dict()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/api/export/excel")
def export_excel():
    global _LAST_REPORT
    if _LAST_REPORT is None:
        audit_mock_project("docx")
    xlsx_bytes = generate_excel_report_bytes(_LAST_REPORT)
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Relatorio_Auditoria_IPC_Stellantis.xlsx"'},
    )


@app.get("/api/export/pdf")
def export_pdf():
    global _LAST_REPORT
    if _LAST_REPORT is None:
        audit_mock_project("docx")
    pdf_bytes = generate_pdf_report_bytes(_LAST_REPORT)
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": 'attachment; filename="Relatorio_Auditoria_IPC_Stellantis.pdf"'},
    )
