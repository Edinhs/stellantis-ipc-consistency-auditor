"""
CLI de execução rápida da Auditoria Bidirecional de Consistência IPC (Stellantis).
Carrega o pacote J3U - 5870 (ou qualquer diretório passado via argumento),
imprime o diagnóstico completo no terminal e salva os relatórios em Excel (.xlsx) e PDF (.pdf).
"""

import sys
from pathlib import Path

from app.parsers import discover_and_parse_project
from app.auditor import run_bidirectional_audit
from app.exporter import generate_excel_report_bytes, generate_pdf_report_bytes
from scripts.generate_mock_data import MOCK_ROOT, generate_all_mock_files


def main():
    target_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else MOCK_ROOT
    if target_dir == MOCK_ROOT and not MOCK_ROOT.exists():
        generate_all_mock_files()

    parsed = discover_and_parse_project(target_dir, prefer_format="docx")
    report = run_bidirectional_audit(
        vfs=parsed["vfs"],
        indications_rows=parsed["indications_rows"],
        parameters_rows=parsed["parameters_rows"],
        project_name="J3U - 5870 | IPC - Instrument Panel Cluster 10 (CAN)",
        files_info=parsed,
    )

    reports_dir = Path(__file__).resolve().parent / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = reports_dir / "Relatorio_Auditoria_IPC_J3U.xlsx"
    pdf_path = reports_dir / "Relatorio_Auditoria_IPC_J3U.pdf"

    xlsx_path.write_bytes(generate_excel_report_bytes(report))
    pdf_path.write_bytes(generate_pdf_report_bytes(report))

    k = report.summary_kpis
    print("=" * 88)
    print("STELLANTIS - AUDITORIA BIDIRECIONAL DE CONSISTÊNCIA IPC (PREVENÇÃO DE RARs)")
    print("=" * 88)
    print(f"Projeto:                  {report.project_name}")
    print(f"VFs Analisadas:           {k['total_vfs_analyzed']} ({', '.join(parsed['vf_files'])})")
    print(f"Planilha Indications:     {parsed['indications_file']} ({k['total_sheet_indications']} linhas)")
    print(f"Planilha Parameters:      {parsed['parameters_file']} ({k['total_sheet_params']} linhas)")
    print(f"Taxa de Conformidade:     {k['conformity_rate_pct']}%")
    print(f"Riscos Únicos de RAR:     {k['total_unique_rar_risks']}")
    print("-" * 88)
    print(f"[REGRA 1: VF <-> Indications] Total de Inconsistências: {k['rule1_unique_issues']}")
    for f in report.rule1_findings:
        if f.status != "OK":
            print(f"  * [{f.direction}] [{f.status}] {f.item_name} (VF Doc: {f.vf_id} | Planilha: {f.sheet_vf_ref})")
            print(f"      -> {f.details}")
    print("-" * 88)
    print(f"[REGRA 2: VF <-> Parameters Table (IPC + SUPPLIER EOL)] Total de Inconsistências: {k['rule2_unique_issues']}")
    for f in report.rule2_findings:
        if f.status != "OK":
            print(f"  * [{f.direction}] [{f.status}] {f.item_name} (VF Doc: {f.vf_id} | Ref: {f.sheet_vf_ref})")
            print(f"      -> {f.details}")
    print("-" * 88)
    print(f"Relatório Excel salvo em: {xlsx_path}")
    print(f"Relatório PDF salvo em:   {pdf_path}")
    print("=" * 88)


if __name__ == "__main__":
    main()
