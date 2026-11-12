"""
Gerador de Pacote de Especificações em Escala Realista (1.215 páginas / 45 VFs)
Reproduz o volume real de um pacote VFS_IPC da Stellantis (1.000+ páginas exportadas do DOORS):
- 45 Vehicle Functions completas (VF101 a VF245)
- Cada VF possui 27 páginas no PDF consolidado (totalizando 1.215 páginas em VFS_IPC_PJ3U_2A_1215P.pdf)
- Cada VF contém múltiplos capítulos e tabelas distratoras reais (CAN Signals, State Machine, DTC Table,
  dezenas de requisitos [REQ_VFxxx_xxxx]) além das seções-alvo:
  * x.12 Indication Management (com texto corrido DOORS + tabela de indicações)
  * x.14.1 Configuration Parameters Table (com parâmetros IPC+SUPPLIER EOL_2, PROXI e outras ECUs, além de VFs com tabela vazia)
"""

import time
from pathlib import Path
import docx
from docx.shared import Pt, RGBColor
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import fitz  # PyMuPDF


BASE_DIR = Path(__file__).resolve().parent.parent
LARGE_ROOT = (
    BASE_DIR
    / "mock_data_large_1200p"
    / "J3U - 5870_FULL_1215_PAGES"
    / "J3U_Work in progress"
    / "03-EE System Requirements"
    / "IPC - Instrument Panel Cluster 10 (CAN)"
)

LARGE_VF_DIR = LARGE_ROOT / "VF-Functional Requirements"
LARGE_IND_DIR = (
    LARGE_ROOT
    / "TS-Technical Requirements"
    / "STS-System Technical Specification"
    / "Section C - HMI Requirements"
)
LARGE_PAR_DIR = (
    LARGE_ROOT
    / "TS-Technical Requirements"
    / "STS-System Technical Specification"
    / "Section B - Software Requirements"
)


# Catálogo de 45 VFs automotivas realistas do cluster IPC Stellantis
AUTOMOTIVE_DOMAINS = [
    ("VF101", "Vehicle Speedometer Display Management", "SPEEDOMETER", "V_SPEED"),
    ("VF102", "Tachometer and Engine RPM Management", "TACHOMETER", "N_RPM"),
    ("VF103", "Fuel Level Gauge and Low Fuel Warning", "FUEL LEVEL", "L_FUEL"),
    ("VF104", "Engine Coolant Temperature Gauge", "COOLANT TEMP", "T_COOLANT"),
    ("VF105", "Total and Partial Odometer Management", "ODOMETER", "D_ODO"),
    ("VF106", "Trip Computer A and B Calculation", "TRIP COMPUTER", "T_TRIP"),
    ("VF107", "Gear Shift Indicator (GSI) Management", "GEAR SHIFT", "N_GSI"),
    ("VF108", "Automatic Transmission Gear Display", "AT GEAR", "T_AT_DISPLAY"),
    ("VF109", "Stop and Start System HMI Management", "STOP START", "T_SS_DELAY"),
    ("VF110", "Cruise Control (CC) HMI Indication", "CRUISE CONTROL", "V_CC_MIN"),
    ("VF111", "Speed Limiter (SL) Visual and Acoustic Warning", "SPEED LIMITER", "V_SL_HYST"),
    ("VF112", "Adaptive Cruise Control (ACC) Cluster HMI", "ACC SYSTEM", "D_ACC_HEADWAY"),
    ("VF113", "Lane Departure Warning (LDW) & LKA HMI", "LANE DEPARTURE", "V_LDW_ACT"),
    ("VF114", "Autonomous Emergency Braking (AEB) Warning", "AEB COLLISION", "T_AEB_POPUP"),
    ("VF115", "Blind Spot Monitoring (BSM) Indication", "BLIND SPOT", "V_BSM_MIN"),
    ("VF116", "Driver Drowsiness Detection (DDD) Coffee Cup", "DRIVER DROWSINESS", "T_DDD_WARN"),
    ("VF117", "Traffic Sign Recognition (TSR) Display", "TRAFFIC SIGN", "T_TSR_HOLD"),
    ("VF118", "Seat Belt Reminder (SBR) Front & Rear", "SEAT BELT REMINDER", "V_SBR_TH"),
    ("VF120", "Doors, Hood and Trunk Open Status Graphic", "DOOR OPEN", "V_DOOR_WARN"),
    ("VF122", "External Lighting & Turn Signals Telltales", "TURN SIGNAL", "F_BLINK_RATE"),
    ("VF124", "High Beam & Auto High Beam Assist HMI", "HIGH BEAM", "V_AHB_DEACT"),
    ("VF126", "Front and Rear Fog Lights Indication", "FOG LIGHT", "T_FOG_FILTER"),
    ("VF128", "Daytime Running Lights (DRL) Fault Warning", "DRL FAULT", "T_DRL_DEBOUNCE"),
    ("VF130", "Alternator & Battery Charging System Warning", "BATTERY CHARGE", "U_BATT_LOW"),
    ("VF132", "Engine Oil Pressure & Level Warning", "OIL PRESSURE", "T_OIL_DELAY"),
    ("VF134", "MIL Check Engine (OBD) Telltale Management", "CHECK ENGINE MIL", "T_MIL_BULBCHK"),
    ("VF136", "Diesel Particulate Filter (DPF) & AdBlue HMI", "DPF REGENERATION", "L_ADBLUE_MIN"),
    ("VF138", "Glow Plug Pre-Heating & Water in Fuel", "GLOW PLUG", "T_GLOW_MAX"),
    ("VF139", "Airbag & Pretensioner Warning Management", "AIR BAG", "tREC"),
    ("VF141", "Electric Power Steering (EPS) Fault Indication", "EPS STEERING", "T_EPS_FILTER"),
    ("VF143", "Hand Brake & Electric Park Brake (EPB) HMI", "HAND BRAKE", "V_TH_HB"),
    ("VF145", "Low Brake Fluid & EBD Brake System Warning", "LOW BRAKE FLUID", "T_LBF_FILTER"),
    ("VF147", "ABS & ESC Stability Control Telltale Logic", "ESC STABILITY", "F_ESC_BLINK"),
    ("VF149", "Hill Holder & Hill Descent Control (HDC) HMI", "HILL HOLDER", "V_HDC_MAX"),
    ("VF150", "Wiper, Washer Fluid & Rain Sensor HMI", "WASHER FLUID", "T_WASHER_FILT"),
    ("VF155", "ParkSense Front & Rear Parking Sensors Graphic", "PARKSENSE", "D_PARK_BEEP"),
    ("VF160", "Tire Pressure Monitoring System (TPMS)", "TPMS PRESSURE", "T_TPMS_BLINK"),
    ("VF165", "Drive Mode Selector (Sport / Eco / Auto) HMI", "DRIVE MODE", "T_MODE_BANNER"),
    ("VF170", "Service Interval & Oil Life Reset Management", "SERVICE DUE", "D_SERVICE_WARN"),
    ("VF175", "Outside Temperature & Ice Warning Snowflake", "ICE WARNING", "T_ICE_TH"),
    ("VF180", "Immobilizer & Passive Entry Key Not Detected", "KEY NOT DETECTED", "T_IMMO_POPUP"),
    ("VF185", "Acoustic Chime Priority & Arbitration Manager", "CHIME ARBITRATION", "T_CHIME_GAP"),
    ("VF190", "Display Brightness & Day/Night Dimming", "DISPLAY DIMMING", "PWM_DAY_DEFAULT"),
    ("VF195", "Cluster Self-Test & Bulb Check Sequence", "BULB CHECK", "T_BULB_CHECK"),
    ("VF200", "CAN-FD Network Timeout & Node Recovery HMI", "NETWORK RECOVERY", "T_NODE_TIMEOUT"),
]


def build_45_vfs_catalog():
    """
    Constrói a estrutura completa de 45 VFs, suas indicações e parâmetros, além das planilhas
    Indications e Parameters Table com inconsistências controladas (RARs) em meio a mais de 200 itens válidos.
    """
    vfs_list = []
    sheet_indications = []
    sheet_parameters = []

    for idx, (vf_id, title_desc, ind_prefix, par_prefix) in enumerate(AUTOMOTIVE_DOMAINS, start=10):
        vf_full_code = f"{vf_id}_V1_PJ3U"
        # Cada VF terá 3 indicações na seção Indication Management (45 * 3 = 135 indicações totais nas VFs)
        ind1_raw = f"{ind_prefix} ACTIVE.V1"
        ind2_raw = f"{ind_prefix} FAIL.V1"
        ind3_raw = f"{ind_prefix} RECOVERY.V1"

        indications = [
            {
                "name": ind1_raw,
                "condition": f"When {par_prefix}_Sts == ACTIVE on CAN-FD network",
                "popup": f"No (Standard {ind_prefix} Telltale)",
                "telltale": f"{ind_prefix} Green/Amber Telltale ON",
            },
            {
                "name": ind2_raw,
                "condition": f"When {par_prefix}_FailSts == PRESENT",
                "popup": f"Yes - Check {title_desc}",
                "telltale": f"{ind_prefix} Warning Telltale Blinking 2Hz",
            },
            {
                "name": ind3_raw,
                "condition": f"When CAN node timeout > {par_prefix}_TIMEOUT",
                "popup": f"Yes - {ind_prefix} System Recovery",
                "telltale": f"{ind_prefix} Telltale Steady ON",
            },
        ]

        # Inserir as indicações na planilha Indications (com algumas inconsistências intencionais realistas):
        # 1. Nas VFs VF114, VF139 e VF143, omitimos a 3ª indicação na planilha (3 MISSING_IN_SHEET)
        # 2. Nas VFs VF124 e VF150, colocamos código de VF divergente na planilha para a 2ª indicação (2 VF_MISMATCH)
        for i_pos, ind_obj in enumerate(indications):
            sheet_name_space = ind_obj["name"].replace(".V1", " V1")
            if i_pos == 2 and vf_id in ("VF114", "VF139", "VF143"):
                # Omitida intencionalmente da planilha Indications -> MISSING_IN_SHEET
                continue
            sheet_vf_col = vf_full_code
            if i_pos == 1 and vf_id in ("VF124", "VF150"):
                sheet_vf_col = "VF999_V1_LEGACY"  # Divergência intencional -> VF_MISMATCH

            sheet_indications.append(
                {
                    "ipc": "HIGH V10",
                    "name": sheet_name_space,
                    "vfs": sheet_vf_col,
                    "hard_tt": f"TT_{vf_id}_{i_pos+1}",
                    "soft_tt": "-",
                    "driven": "CAN",
                    "std_behav": ind_obj["telltale"],
                    "std_par": "Continuous",
                    "rec_behav": "ON",
                    "display_msg": ind_obj["popup"],
                    "display_rec": "OFF",
                }
            )

        # Parâmetros da VF:
        # Algumas VFs (a cada 9 VFs, ex: VF150) têm Configuration Parameters Table vazia (como mostrado no vídeo aos 18:22)
        parameters = []
        if vf_id not in ("VF108", "VF126", "VF150", "VF180", "VF195"):
            par1_name = f"{par_prefix}_TH"
            par2_name = f"{par_prefix}_DEBOUNCE"
            par_proxi = f"CFG_{vf_id}_PROXI_EN"
            par_other_ecu = f"T_BCM_{vf_id}_RAW"

            parameters = [
                {
                    "id": f"PAR_{vf_id}_01",
                    "name": par1_name,
                    "first_trial": "10,5",
                    "range": "1,0÷50,0",
                    "resolution": "0,5",
                    "unit": "s",
                    "component": "IPC",
                    "implementation": "SUPPLIER EOL_2",
                    "xcp": "NO",
                    "impact": "HMI Calibration",
                    "description": f"Primary threshold parameter for {title_desc} on IPC.",
                },
                {
                    "id": f"PAR_{vf_id}_02",
                    "name": par2_name,
                    "first_trial": "2,0",
                    "range": "0,5÷10,0",
                    "resolution": "0,1",
                    "unit": "s",
                    "component": "IPC",
                    "implementation": "SUPPLIER EOL_2",
                    "xcp": "NO",
                    "impact": "Signal Filtering",
                    "description": f"Debounce timer for {title_desc} warning activation.",
                },
                {
                    "id": f"PAR_{vf_id}_03",
                    "name": par_proxi,
                    "first_trial": "1",
                    "range": "0÷1",
                    "resolution": "1",
                    "unit": "-",
                    "component": "IPC",
                    "implementation": "PROXI",  # Deve ser segregado pelo filtro!
                    "xcp": "NO",
                    "impact": "Variant Coding",
                    "description": f"PROXI enable bit for {title_desc}.",
                },
                {
                    "id": f"PAR_{vf_id}_04",
                    "name": par_other_ecu,
                    "first_trial": "100",
                    "range": "10÷500",
                    "resolution": "10",
                    "unit": "ms",
                    "component": "BCM",  # Outra ECU -> Deve ser segregado pelo filtro!
                    "implementation": "SUPPLIER EOL",
                    "xcp": "YES",
                    "impact": "Input Acquisition",
                    "description": f"BCM internal acquisition timer for {title_desc}.",
                },
            ]

            # Popular a planilha IPC_Parameters_Table com inconsistências controladas:
            # 1. Nas VFs VF118, VF143 e VF165, omitimos o 2º parâmetro EOL na planilha (3 MISSING_IN_SHEET)
            # 2. Nas VFs VF130 e VF160, deixamos First Trial Value em branco na planilha para o 1º parâmetro (2 VALUE_MISMATCH_OR_EMPTY)
            for p_pos, p_obj in enumerate([parameters[0], parameters[1]]):
                if p_pos == 1 and vf_id in ("VF118", "VF143", "VF165"):
                    continue  # Omitido da planilha -> MISSING_IN_SHEET
                ftv_sheet = p_obj["first_trial"]
                if p_pos == 0 and vf_id in ("VF130", "VF160"):
                    ftv_sheet = ""  # Em branco na planilha -> VALUE_MISMATCH_OR_EMPTY

                sheet_parameters.append(
                    {
                        "name": p_obj["name"],
                        "first_trial": ftv_sheet,
                        "range": p_obj["range"],
                        "resolution": p_obj["resolution"],
                        "unit": p_obj["unit"],
                        "description": p_obj["description"],
                        "reference": vf_id,
                        "p363": "x",
                        "p376": "x",
                        "pj3u": "x",
                        "hw_col": "ALL",
                    }
                )

        vfs_list.append(
            {
                "vf_id": vf_id,
                "vf_full_code": vf_full_code,
                "title": f"{vf_id} - {title_desc}",
                "section_prefix": str(idx),
                "indications": indications,
                "parameters": parameters,
            }
        )

    # Adicionar 3 Indicações ÓRFÃS na planilha Indications (ORPHAN_IN_SHEET)
    for orph_idx, (orph_name, orph_vf) in enumerate(
        [
            ("NIGHT VISION PEDESTRIAN ALERT V1", "VF888_V1_P363"),
            ("AIR SUSPENSION LEVEL FAULT V1", "VF889_V1_P376"),
            ("TRAILER SWAY DAMPING ACTIVE V1", ""),
        ],
        start=1,
    ):
        sheet_indications.append(
            {
                "ipc": "HIGH V10",
                "name": orph_name,
                "vfs": orph_vf,
                "hard_tt": f"TT_ORPH_{orph_idx}",
                "soft_tt": "-",
                "driven": "CAN",
                "std_behav": "Legacy Carryover",
                "std_par": "-",
                "rec_behav": "OFF",
                "display_msg": "Orphan indication from legacy platform",
                "display_rec": "OFF",
            }
        )

    # Adicionar 3 Parâmetros ÓRFÃOS na planilha IPC_Parameters_Table (ORPHAN_IN_SHEET)
    for p_orph_name, p_orph_ref in [
        ("T_LEGACY_P363_COMPASS_CAL", "Legacy P363"),
        ("V_GCC_OVERSPEED_BUZZER_TH", "VF777 (Middle East Only)"),
        ("K_ANALOG_GAUGE_STEPPER_GAIN", "Old Analog Cluster"),
    ]:
        sheet_parameters.append(
            {
                "name": p_orph_name,
                "first_trial": "15,0",
                "range": "0÷100",
                "resolution": "1",
                "unit": "s",
                "description": "Orphan parameter left over from previous vehicle program.",
                "reference": p_orph_ref,
                "p363": "x",
                "p376": "",
                "pj3u": "x",
                "hw_col": "ALL",
            }
        )

    return vfs_list, sheet_indications, sheet_parameters


def generate_1215_page_pdf(pdf_path: Path, vfs_list: list[dict], pages_per_vf: int = 27) -> int:
    """
    Gera um arquivo PDF consolidado realista com 1.215 páginas (45 VFs x 27 páginas por VF),
    contendo todas as seções DOORS (arquitetura, sinais CAN, máquina de estados, dezenas de páginas de
    requisitos funcionais [REQ_VFxxx_xxxx], seção Indication Management, diagnóstico DTC e Configuration Parameters Table).
    """
    pdf_doc = fitz.open()
    global_page_num = 0

    for vf in vfs_list:
        sec = vf["section_prefix"]
        vf_id = vf["vf_id"]
        vf_code = vf["vf_full_code"]

        for local_p in range(1, pages_per_vf + 1):
            global_page_num += 1
            page = pdf_doc.new_page(width=595, height=842)  # A4
            y = 38

            # Cabeçalho de página DOORS / Stellantis
            page.insert_text(
                (40, y),
                f"STELLANTIS - EE SYSTEM REQUIREMENTS | VFS_IPC_PJ3U_2A (FULL 1200+ PAGES) | Page {global_page_num}",
                fontsize=8,
                color=(0.35, 0.35, 0.35),
            )
            y += 20

            if local_p == 1:
                # Página 1 da VF: Título, Metadados, Escopo e Tabela Distratora de Sinais CAN
                page.insert_text((40, y), f"{sec} {vf['title']} ({vf_code})", fontsize=13, color=(0.08, 0.18, 0.38))
                y += 18
                page.insert_text(
                    (40, y),
                    f"Document Code: {vf_code} | Target Subsystem: IPC - Instrument Panel Cluster 10 (CAN)",
                    fontsize=9,
                    color=(0.2, 0.2, 0.2),
                )
                y += 22
                page.insert_text((40, y), f"{sec}.1 Purpose and Functional Overview", fontsize=10.5, color=(0.1, 0.2, 0.4))
                y += 16
                page.insert_text(
                    (45, y),
                    f"[REQ_{vf_id}_0001] The IPC shall manage the visual, acoustic and diagnostic behavior for {vf['title']}.",
                    fontsize=8.5,
                )
                y += 22
                page.insert_text((40, y), f"{sec}.4 Input / Output CAN-FD Signals Table (Distractor Table)", fontsize=10.5)
                y += 15
                page.insert_text((45, y), "CAN Message | Signal Name | Network | Period (ms) | Sender -> Receiver", fontsize=8)
                y += 13
                for sig_i in range(1, 18):
                    page.insert_text(
                        (45, y),
                        f"CAN_ROW | STATUS_B_CAN_{sig_i:02d} | {vf_id}_SigSts_{sig_i:02d} | CAN_FD_BH | 100 ms | BCM -> IPC",
                        fontsize=7.8,
                    )
                    y += 13

            elif local_p == 22:
                # Página 22 da VF: Seção x.12 Indication Management (Texto DOORS + Tabela de Indicações)
                page.insert_text((40, y), f"{sec}.12 Indication Management", fontsize=11, color=(0.08, 0.18, 0.38))
                y += 18
                for ind in vf["indications"]:
                    # Parágrafo em texto corrido estilo DOORS (sem usar IND_ROW) na 1ª indicação para testar leitura híbrida!
                    page.insert_text(
                        (45, y),
                        f"[REQ_{vf_id}_IND] {ind['condition']}, the IPC cluster shall activate {ind['name']} on the display.",
                        fontsize=8.5,
                    )
                    y += 14
                    line = f"IND_ROW | {ind['name']} | {ind['condition']} | {ind['popup']} | {ind['telltale']}"
                    page.insert_text((45, y), line[:110], fontsize=8.2)
                    y += 15

            elif local_p == 26:
                # Página 26 da VF: Seção x.14.1 Configuration Parameters Table
                page.insert_text((40, y), f"{sec}.14 Configuration Parameters", fontsize=11, color=(0.08, 0.18, 0.38))
                y += 16
                page.insert_text((40, y), f"{sec}.14.1 Configuration Parameters Table", fontsize=10.5, color=(0.08, 0.18, 0.38))
                y += 16
                page.insert_text(
                    (45, y),
                    "ID | Parameter Name | First Trial Value | Range Value | Resolution | Unit | Component | Implementation | XCP | Impact | Description",
                    fontsize=7.5,
                )
                y += 14
                if vf["parameters"]:
                    for par in vf["parameters"]:
                        pline = (
                            f"PAR_ROW | {par['id']} | {par['name']} | {par['first_trial']} | {par['range']} | "
                            f"{par['resolution']} | {par['unit']} | {par['component']} | {par['implementation']} | "
                            f"{par['xcp']} | {par['impact']} | {par['description']}"
                        )
                        page.insert_text((45, y), pline[:115], fontsize=7.5)
                        y += 14
                else:
                    page.insert_text((45, y), "PAR_ROW | - | - | - | - | - | - | - | - | - | - | -", fontsize=7.5)

            else:
                # Páginas intermediárias (2..21, 23..25, 27): Requisitos DOORS extensos, máquinas de estado e tabelas DTC
                sub_num = local_p if local_p < 12 else local_p + 2
                page.insert_text(
                    (40, y),
                    f"{sec}.{sub_num} Detailed Functional Requirements & State Logic — {vf_id} (Part {local_p}/{pages_per_vf})",
                    fontsize=10,
                    color=(0.15, 0.25, 0.45),
                )
                y += 18
                for req_line in range(1, 28):
                    req_id = f"REQ_{vf_id}_{local_p:02d}{req_line:02d}"
                    page.insert_text(
                        (45, y),
                        f"[{req_id}] When KeySts == KEY_ON and network voltage is within [9.0V .. 16.0V], subsystem state machine step {req_line} executes.",
                        fontsize=7.8,
                        color=(0.15, 0.15, 0.15),
                    )
                    y += 16

    pdf_doc.save(str(pdf_path))
    total_pages = len(pdf_doc)
    pdf_doc.close()
    return total_pages


def generate_consolidated_45vfs_docx(docx_path: Path, vfs_list: list[dict]):
    """Gera o arquivo Word (.docx) consolidado com todas as 45 VFs, tabelas distratoras CAN/DTC e seções-alvo."""
    doc = docx.Document()
    doc.add_heading("STELLANTIS - COMPLETE IPC VEHICLE FUNCTIONS SPECIFICATION (45 VFs / 1200+ PAGES EQUIVALENT)", 0)

    for vf in vfs_list:
        sec = vf["section_prefix"]
        vf_id = vf["vf_id"]
        doc.add_heading(f"{sec} {vf['title']} ({vf['vf_full_code']})", level=1)
        p_meta = doc.add_paragraph()
        p_meta.add_run("Document Code: ").bold = True
        p_meta.add_run(f"{vf['vf_full_code']}   |   Subsystem: IPC - Instrument Panel Cluster 10 (CAN)")

        # Seção Distratora: Tabela de Sinais CAN (Deve ser ignorada pelo parser)
        doc.add_heading(f"{sec}.4 Input and Output CAN Signals", level=2)
        can_tbl = doc.add_table(rows=1, cols=4)
        for i, h in enumerate(["CAN Message", "Signal Name", "Sender", "Period"]):
            can_tbl.rows[0].cells[i].text = h
        row_c = can_tbl.add_row().cells
        row_c[0].text = f"STATUS_BH_{vf_id}"
        row_c[1].text = f"{vf_id}_SignalSts"
        row_c[2].text = "BCM"
        row_c[3].text = "100 ms"

        # Seção Alvo 1: x.12 Indication Management (com parágrafo DOORS + tabela)
        doc.add_heading(f"{sec}.12 Indication Management", level=2)
        if vf["indications"]:
            first_ind = vf["indications"][0]
            doc.add_paragraph(
                f"[REQ_{vf_id}_IND_01] {first_ind['condition']}, the IPC shall activate the {first_ind['name']} indication."
            )

        ind_tbl = doc.add_table(rows=1, cols=4)
        for i, h in enumerate(["Indication Name", "Activation Condition", "Display Pop-up / Message", "Telltale Behavior"]):
            ind_tbl.rows[0].cells[i].text = h
        for ind in vf["indications"]:
            rc = ind_tbl.add_row().cells
            rc[0].text = ind["name"]
            rc[1].text = ind["condition"]
            rc[2].text = ind["popup"]
            rc[3].text = ind["telltale"]

        # Seção Distratora Pós-Indication: x.13 Diagnostic Fault Management
        doc.add_heading(f"{sec}.13 Diagnostic & Fault Management", level=2)
        doc.add_paragraph(f"[REQ_{vf_id}_DIAG_01] In case of CAN bus-off or node timeout, recovery rules apply.")

        # Seção Alvo 2: x.14.1 Configuration Parameters Table
        doc.add_heading(f"{sec}.14 Configuration Parameters", level=2)
        doc.add_heading(f"{sec}.14.1 Configuration Parameters Table", level=3)
        p_headers = [
            "ID",
            "Parameter Name",
            "First Trial Value",
            "Range Value",
            "Resolution",
            "Unit",
            "Component",
            "Implementation",
            "XCP Availability",
            "Impact",
            "Parameter Description",
        ]
        par_tbl = doc.add_table(rows=1, cols=len(p_headers))
        for i, h in enumerate(p_headers):
            par_tbl.rows[0].cells[i].text = h

        if vf["parameters"]:
            for par in vf["parameters"]:
                rc = par_tbl.add_row().cells
                vals = [
                    par["id"],
                    par["name"],
                    par["first_trial"],
                    par["range"],
                    par["resolution"],
                    par["unit"],
                    par["component"],
                    par["implementation"],
                    par["xcp"],
                    par["impact"],
                    par["description"],
                ]
                for i, v in enumerate(vals):
                    rc[i].text = v
        else:
            rc = par_tbl.add_row().cells
            for i in range(len(p_headers)):
                rc[i].text = "-"

    doc.save(str(docx_path))


def save_indications_xlsx(xlsx_path: Path, rows: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Full TFT 7 Inch. - LATAM"
    headers = [
        "IPC",
        "Name",
        "VFs",
        "Hard TT code",
        "Soft TT code",
        "Driven",
        "Hard TT Std Behav",
        "Hard TT Std Behav Par",
        "Hard TT Rec Behav",
        "Display messages",
        "Display recovery",
    ]
    ws.append(headers)
    for r in rows:
        ws.append(
            [
                r["ipc"],
                r["name"],
                r["vfs"],
                r["hard_tt"],
                r["soft_tt"],
                r["driven"],
                r["std_behav"],
                r["std_par"],
                r["rec_behav"],
                r["display_msg"],
                r["display_rec"],
            ]
        )
    wb.save(str(xlsx_path))


def save_parameters_xlsx(xlsx_path: Path, rows: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Parameters Table"
    headers = [
        "Parameter Name",
        "First Trial Value",
        "Range Value",
        "Resolution",
        "Unit",
        "Parameter Description",
        "Reference",
        "363",
        "376",
        "J3U",
        "Column Used by IPC HW",
    ]
    ws.append(headers)
    for p in rows:
        ws.append(
            [
                p["name"],
                p["first_trial"],
                p["range"],
                p["resolution"],
                p["unit"],
                p["description"],
                p["reference"],
                p["p363"],
                p["p376"],
                p["pj3u"],
                p["hw_col"],
            ]
        )
    wb.save(str(xlsx_path))


def generate_large_scale_dataset() -> dict:
    t0 = time.perf_counter()
    LARGE_VF_DIR.mkdir(parents=True, exist_ok=True)
    LARGE_IND_DIR.mkdir(parents=True, exist_ok=True)
    LARGE_PAR_DIR.mkdir(parents=True, exist_ok=True)

    vfs_list, sheet_indications, sheet_parameters = build_45_vfs_catalog()

    pdf_path = LARGE_VF_DIR / "VFS_IPC_PJ3U_2A_1215_PAGES.pdf"
    total_pdf_pages = generate_1215_page_pdf(pdf_path, vfs_list, pages_per_vf=27)

    docx_path = LARGE_VF_DIR / "VFS_IPC_PJ3U_2A_45VFs.docx"
    generate_consolidated_45vfs_docx(docx_path, vfs_list)

    ind_path = LARGE_IND_DIR / "Indications_J3U_R4_FULL.xlsx"
    save_indications_xlsx(ind_path, sheet_indications)

    par_path = LARGE_PAR_DIR / "IPC_Parameters_Table_PJ3U_R5_FULL.xlsx"
    save_parameters_xlsx(par_path, sheet_parameters)

    elapsed = round(time.perf_counter() - t0, 2)
    return {
        "project_root": str(LARGE_ROOT),
        "pdf_path": str(pdf_path),
        "total_pdf_pages": total_pdf_pages,
        "docx_path": str(docx_path),
        "indications_xlsx": str(ind_path),
        "parameters_xlsx": str(par_path),
        "vfs_count": len(vfs_list),
        "generation_time_sec": elapsed,
    }


if __name__ == "__main__":
    info = generate_large_scale_dataset()
    print("Pacote de Grande Escala (1.215 páginas / 45 VFs) gerado com sucesso:")
    for k, v in info.items():
        print(f"  - {k}: {v}")
