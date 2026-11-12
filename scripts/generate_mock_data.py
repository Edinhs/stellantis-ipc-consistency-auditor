"""
Gerador de dados fictícios realistas (Mock Data) para o projeto Stellantis IPC Consistency Auditor.
Reproduz fielmente a estrutura de diretórios e o layout dos documentos exibidos na reunião:
- VFs em .docx e .pdf (pacote consolidado VFS_IPC_PJ3U_2A e arquivos individuais em IPC_J3U_2A/)
- Planilha de Indicações (Section C - HMI Requirements/Indications_J3U_R4.xlsx)
- Planilha de Parâmetros (Section B - Software Requirements/IPC_Parameters_Table_PJ3U_R5.xlsx)

Inclui casos válidos e inconsistências intencionais (cenários reais de RAR) para validação bidirecional.
"""

from pathlib import Path
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import fitz  # PyMuPDF


BASE_DIR = Path(__file__).resolve().parent.parent
MOCK_ROOT = (
    BASE_DIR
    / "mock_data"
    / "J3U - 5870"
    / "J3U_Work in progress"
    / "03-EE System Requirements"
    / "IPC - Instrument Panel Cluster 10 (CAN)"
)

VF_DIR = MOCK_ROOT / "VF-Functional Requirements"
VF_INDIVIDUAL_DIR = VF_DIR / "IPC_J3U_2A"
INDICATIONS_DIR = (
    MOCK_ROOT
    / "TS-Technical Requirements"
    / "STS-System Technical Specification"
    / "Section C - HMI Requirements"
)
PARAMETERS_DIR = (
    MOCK_ROOT
    / "TS-Technical Requirements"
    / "STS-System Technical Specification"
    / "Section B - Software Requirements"
)


# Dados das VFs fictícias baseadas fielmente nos exemplos do vídeo (VF139, VF143, VF145, VF150, VF118, VF210)
VFS_DATA = [
    {
        "vf_id": "VF139",
        "vf_full_code": "VF139_V1_PJ3U",
        "title": "VF139 - Airbag Warning Management",
        "filename_stem": "VF139_Airbag_Warning_Management_PJ3U",
        "section_prefix": "38",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "ORC (Airbag ECU)", "BCM"],
        "indications": [
            {
                "name": "AIR BAG FAIL.V1",
                "condition": "When FailPresent == 1 or AirBagLampReq == ON",
                "Popup": "Yes - Airbag failure, see handbook",
                "telltale": "Airbag Warning Red Telltale ON",
            },
            {
                "name": "AIR BAG TEMPORARY CALIBRATION.V1",
                "condition": "When ORC is in Plant Mode / Temporary Calibration active",
                "Popup": "Yes - Airbag Plant Calibration Mode",
                "telltale": "Airbag Warning Blinking (2 Hz)",
            },
            {
                "name": "AIR BAG RECOVERY.V1",
                "condition": "When CAN node ORC_FD_1 is in timeout > tREC",
                "Popup": "Yes - Check Airbag System",
                "telltale": "Airbag Warning ON (Recovery Behavior)",
            },
            {
                # INTENCIONAL RAR #1: Chamada na VF139, mas AUSENTE na planilha Indications_J3U_R4.xlsx
                "name": "PASSENGER AIR BAG OFF.V1",
                "condition": "When PassengerAirBagStatus == DISABLED",
                "Popup": "Yes - Passenger Airbag Deactivated",
                "telltale": "PAB Amber Telltale ON",
            },
        ],
        "parameters": [
            {
                "id": "PAR_139_01",
                "name": "tREC",
                "first_trial": "2,5",
                "range": "0,1÷5",
                "resolution": "0,1",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "Diagnosis",
                "description": "Time-out of the signals received for node recovery (tolerance ± 10%).",
            },
            {
                "id": "PAR_139_02",
                "name": "T_AIRBAG_STARTUP_CHK",
                "first_trial": "4,0",
                "range": "2,0÷6,0",
                "resolution": "0,5",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "HMI",
                "description": "Duration of the Airbag telltale bulb check phase after Key-ON.",
            },
            {
                "id": "PAR_139_03",
                "name": "CFG_PAB_PRESENT",
                "first_trial": "1",
                "range": "0÷1",
                "resolution": "1",
                "unit": "-",
                "component": "IPC",
                "implementation": "PROXI",  # Deve ser filtrado (vai para especificação PROXI, não Parameters Table EOL)
                "xcp": "NO",
                "impact": "Vehicle Configuration",
                "description": "Enables Passenger Airbag inhibition telltale via PROXI byte.",
            },
            {
                "id": "PAR_139_04",
                "name": "T_ORC_CRASH_LATCH",
                "first_trial": "50",
                "range": "10÷200",
                "resolution": "10",
                "unit": "ms",
                "component": "ORC",  # Outro módulo (não IPC) -> Deve ser ignorado pelo filtro IPC
                "implementation": "SUPPLIER EOL",
                "xcp": "YES",
                "impact": "Safety",
                "description": "Crash deployment signal latch time in ORC ECU.",
            },
        ],
    },
    {
        "vf_id": "VF143",
        "vf_full_code": "VF143_V1_PJ3U",
        "title": "VF143 - Hand Brake Management",
        "filename_stem": "VF143_Hand_Brake_Management_PJ3U",
        "section_prefix": "62",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "BCM", "BSM (ESC)"],
        "indications": [
            {
                "name": "HAND BRAKE.V1",
                "condition": "When HandBrakeSwitchSts == ENGAGED and VehicleSpeed <= V_TH_HB",
                "Popup": "No",
                "telltale": "Park Brake Red Telltale ON",
            },
            {
                "name": "HAND BRAKE VEHICLE IN MOVEMENT.V1",
                "condition": "When HandBrakeSwitchSts == ENGAGED and VehicleSpeed > V_TH_HB for t > T_HB_WARN",
                "Popup": "Yes - Release Parking Brake + Acoustic Chime",
                "telltale": "Park Brake Red Telltale Blinking",
            },
            {
                # INTENCIONAL RAR #2: Chamada na VF143, mas AUSENTE na planilha Indications_J3U_R4.xlsx
                "name": "EPB AUTO HOLD ACTIVE.V1",
                "condition": "When EPB_AutoHoldSts == ACTIVE",
                "Popup": "Yes - Auto Hold Active",
                "telltale": "Auto Hold Green Telltale ON",
            },
        ],
        "parameters": [
            {
                "id": "PAR_143_01",
                "name": "V_TH_HB",
                "first_trial": "5",
                "range": "3÷15",
                "resolution": "1",
                "unit": "km/h",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "HMI Warning",
                "description": "Vehicle speed threshold to trigger Hand Brake engaged with vehicle in movement warning.",
            },
            {
                "id": "PAR_143_02",
                "name": "T_HB_WARN",
                "first_trial": "1,5",
                "range": "0,5÷5,0",
                "resolution": "0,5",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "HMI Warning",
                "description": "Debounce time with vehicle speed > V_TH_HB before acoustic warning activation.",
            },
            {
                # INTENCIONAL RAR #3: Parâmetro IPC + SUPPLIER EOL_2 presente na VF143, mas AUSENTE na Parameters Table!
                "id": "PAR_143_03",
                "name": "T_HB_CHIME_REPEAT",
                "first_trial": "30",
                "range": "10÷120",
                "resolution": "5",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "Acoustic Chime",
                "description": "Repetition cycle period for Hand Brake acoustic warning while vehicle remains in motion.",
            },
            {
                "id": "PAR_143_04",
                "name": "T_BCM_HB_DEBOUNCE",
                "first_trial": "100",
                "range": "20÷500",
                "resolution": "10",
                "unit": "ms",
                "component": "BCM",  # Outro componente -> filtrado
                "implementation": "SUPPLIER EOL",
                "xcp": "NO",
                "impact": "Input Acquisition",
                "description": "Hardwired handbrake switch debounce time inside Body Computer Module.",
            },
        ],
    },
    {
        "vf_id": "VF145",
        "vf_full_code": "VF145_V1_PJ3U",
        "title": "VF145 - Low Brake Fluid Level Management",
        "filename_stem": "VF145_Low_Brake_Fluid_Level_Management_PJ3U",
        "section_prefix": "64",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "BCM"],
        "indications": [
            {
                "name": "LOW BRAKE FLUID.V1",
                "condition": "When BrakeFluidLevelSts == LOW",
                "Popup": "Yes - Brake fluid level low, stop vehicle safely",
                "telltale": "Brake System Red Telltale ON",
            },
            {
                "name": "BRAKE SYSTEM FAIL.V1",
                "condition": "When EBD_FailSts == PRESENT",
                "Popup": "Yes - Brake system failure",
                "telltale": "Brake System Red + ABS Amber Telltales ON",
            },
        ],
        "parameters": [
            {
                "id": "PAR_145_01",
                "name": "T_LBF_FILTER",
                "first_trial": "2,0",
                "range": "0,5÷5,0",
                "resolution": "0,1",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "HMI Filter",
                "description": "Filtering time on CAN signal BrakeFluidLevelSts before turning ON red telltale.",
            },
        ],
    },
    {
        "vf_id": "VF150",
        "vf_full_code": "VF150_V2_PJ3U",
        "title": "VF150 - Wiper and Washer Management",
        "filename_stem": "VF150_Wiper_and_Washer_Management_PJ3U",
        "section_prefix": "71",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "BCM", "RLS"],
        "indications": [
            {
                # INTENCIONAL RAR #4: Na VF150 é chamada, mas na planilha Indications está associada a VF149_V1_PJ3U (VF Mismatch)
                "name": "WINDSHIELD WASHER FLUID LOW.V1",
                "condition": "When WasherFluidLevelSts == LOW",
                "Popup": "Yes - Top up windshield washer fluid",
                "telltale": "Washer Fluid Amber Icon ON",
            },
            {
                "name": "RAIN SENSOR FAIL.V1",
                "condition": "When RainSensorFailSts == FAIL",
                "Popup": "Yes - Automatic wipers unavailable",
                "telltale": "Dedicated Popup Icon",
            },
        ],
        "parameters": [],  # Exemplo real citado pelo Gabriel no vídeo (18:22): VF com Configuration Parameters Table vazia!
    },
    {
        "vf_id": "VF118",
        "vf_full_code": "VF118_V3_PJ3U",
        "title": "VF118 - Seat Belt Reminder (SBR) Management",
        "filename_stem": "VF118_Seat_Belt_Reminder_Management_PJ3U",
        "section_prefix": "29",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "ORC", "BCM"],
        "indications": [
            {
                "name": "SEAT BELT REMINDER FRONT.V1",
                "condition": "When FrontBeltSts == UNBUCKLED",
                "Popup": "No (Dedicated SBR Graphic)",
                "telltale": "Seat Belt Red Telltale ON / Blinking",
            },
            {
                "name": "SEAT BELT REMINDER REAR.V1",
                "condition": "When RearBeltBuckleChange == UNBUCKLED_IN_MOTION",
                "Popup": "Yes - Rear Seat Belt Status Graphic",
                "telltale": "Rear SBR Status Icons (3 seats)",
            },
        ],
        "parameters": [
            {
                "id": "PAR_118_01",
                "name": "V_SBR_SPEED_TH",
                "first_trial": "20",
                "range": "10÷30",
                "resolution": "1",
                "unit": "km/h",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "ECE-R16 Compliance",
                "description": "Vehicle speed threshold for Seat Belt Reminder second stage visual and audible alarm.",
            },
            {
                "id": "PAR_118_02",
                "name": "T_SBR_CHIME_DURATION",
                "first_trial": "90",
                "range": "30÷180",
                "resolution": "5",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "ECE-R16 Compliance",
                "description": "Total duration of SBR acoustic chime sequence before automatic muting.",
            },
            {
                # INTENCIONAL RAR #5: Parâmetro IPC + SUPPLIER EOL na VF118, mas AUSENTE na Parameters Table!
                "id": "PAR_118_03",
                "name": "D_SBR_DIST_TH",
                "first_trial": "500",
                "range": "100÷1000",
                "resolution": "50",
                "unit": "m",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "ECE-R16 Compliance",
                "description": "Distance threshold travelled with unbuckled belt to trigger SBR alarm.",
            },
            {
                "id": "PAR_118_04",
                "name": "CFG_REAR_SBR_ENABLE",
                "first_trial": "1",
                "range": "0÷1",
                "resolution": "1",
                "unit": "-",
                "component": "IPC",
                "implementation": "PROXI",  # PROXI -> Filtrado
                "xcp": "NO",
                "impact": "Market Variant",
                "description": "PROXI configuration bit to enable/disable Rear Seat Belt Reminder visualization.",
            },
        ],
    },
    {
        "vf_id": "VF210",
        "vf_full_code": "VF210_V1_PJ3U",
        "title": "VF210 - Tire Pressure Monitoring System (TPMS)",
        "filename_stem": "VF210_Tire_Pressure_Monitoring_System_PJ3U",
        "section_prefix": "88",
        "author": "STELLANTIS EE SYSTEM TEAM",
        "components_involved": ["IPC", "BSM", "RFHM"],
        "indications": [
            {
                "name": "TPMS LOW PRESSURE.V1",
                "condition": "When TirePressureWarn == LOW_PRESSURE",
                "Popup": "Yes - Check tire pressure",
                "telltale": "TPMS Amber Telltale ON",
            },
            {
                "name": "TPMS SYSTEM FAIL.V1",
                "condition": "When TPMS_SystemSts == FAULT",
                "Popup": "Yes - TPMS unavailable, service required",
                "telltale": "TPMS Amber Telltale Blinking 75s then Steady ON",
            },
        ],
        "parameters": [
            {
                # INTENCIONAL RAR #6: Existe na VF210 com First Trial Value = 75, mas na Parameters Table está com First Trial Value em BRANCO ("-")
                "id": "PAR_210_01",
                "name": "T_TPMS_BLINK_TIME",
                "first_trial": "75",
                "range": "60÷90",
                "resolution": "5",
                "unit": "s",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "FMVSS138 / Regulation",
                "description": "Blinking duration of TPMS malfunction telltale before remaining continuously illuminated.",
            },
            {
                "id": "PAR_210_02",
                "name": "P_TPMS_WARN_HYST",
                "first_trial": "0,2",
                "range": "0,1÷0,5",
                "resolution": "0,05",
                "unit": "bar",
                "component": "IPC",
                "implementation": "SUPPLIER EOL_2",
                "xcp": "NO",
                "impact": "HMI Display",
                "description": "Pressure hysteresis to clear low tire pressure warning on IPC display.",
            },
        ],
    },
]


# Linhas da planilha Section C - HMI Requirements / Indications_J3U_R4.xlsx
# Observação: A planilha usa espaço antes da versão ("HAND BRAKE V1") enquanto o texto da VF usa ponto ("HAND BRAKE.V1"),
# exatamente como demonstrado pelo Gabriel no vídeo (09:50 - 11:30).
INDICATIONS_EXCEL_ROWS = [
    # VF139 (Nota: PASSENGER AIR BAG OFF V1 foi omitida intencionalmente para simular RAR de indicação ausente!)
    {
        "ipc": "HIGH V10",
        "name": "AIR BAG FAIL V1",
        "vfs": "VF139_V1_PJ3U",
        "hard_tt": "TT_AIRBAG_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady",
        "std_par": "Continuous",
        "rec_behav": "ON",
        "display_msg": "MSG_014 - Airbag Failure",
        "display_rec": "No recovery msg",
    },
    {
        "ipc": "HIGH V10",
        "name": "AIR BAG TEMPORARY CALIBRATION V1",
        "vfs": "VF139_V1_PJ3U",
        "hard_tt": "TT_AIRBAG_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "Blinking 2Hz",
        "std_par": "Duty 50%",
        "rec_behav": "OFF",
        "display_msg": "MSG_015 - Plant Calibration",
        "display_rec": "OFF",
    },
    {
        "ipc": "HIGH V10",
        "name": "AIR BAG RECOVERY V1",
        "vfs": "VF139_V1_PJ3U",
        "hard_tt": "TT_AIRBAG_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady (Node Timeout)",
        "std_par": "t > tREC",
        "rec_behav": "ON",
        "display_msg": "MSG_016 - Check Airbag System",
        "display_rec": "ON",
    },
    # VF143 (Nota: EPB AUTO HOLD ACTIVE V1 foi omitida intencionalmente para simular RAR!)
    {
        "ipc": "HIGH V10",
        "name": "HAND BRAKE V1",
        "vfs": "VF143_V1_PJ3U",
        "hard_tt": "TT_BRAKE_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady",
        "std_par": "Speed <= V_TH_HB",
        "rec_behav": "OFF",
        "display_msg": "None",
        "display_rec": "OFF",
    },
    {
        "ipc": "HIGH V10",
        "name": "HAND BRAKE VEHICLE IN MOVEMENT V1",
        "vfs": "VF143_V1_PJ3U",
        "hard_tt": "TT_BRAKE_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "Blinking 1Hz + Chime",
        "std_par": "Speed > V_TH_HB",
        "rec_behav": "OFF",
        "display_msg": "MSG_042 - Release Parking Brake",
        "display_rec": "OFF",
    },
    # VF145
    {
        "ipc": "HIGH V10",
        "name": "LOW BRAKE FLUID V1",
        "vfs": "VF145_V1_PJ3U",
        "hard_tt": "TT_BRAKE_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady + Chime",
        "std_par": "t > T_LBF_FILTER",
        "rec_behav": "ON",
        "display_msg": "MSG_045 - Brake Fluid Low",
        "display_rec": "ON",
    },
    {
        "ipc": "HIGH V10",
        "name": "BRAKE SYSTEM FAIL V1",
        "vfs": "VF145_V1_PJ3U",
        "hard_tt": "TT_BRAKE_RED",
        "soft_tt": "TT_ABS_AMBER",
        "driven": "CAN",
        "std_behav": "ON Steady",
        "std_par": "Immediate",
        "rec_behav": "ON",
        "display_msg": "MSG_046 - Brake System Fail",
        "display_rec": "ON",
    },
    # VF150 (Nota: WINDSHIELD WASHER FLUID LOW V1 está com VF errada VF149_V1_PJ3U em vez de VF150!)
    {
        "ipc": "HIGH V10",
        "name": "WINDSHIELD WASHER FLUID LOW V1",
        "vfs": "VF149_V1_PJ3U",  # INTENCIONAL RAR #4: Divergência de código de VF na planilha Indications
        "hard_tt": "-",
        "soft_tt": "SOFT_WASHER_AMBER",
        "driven": "CAN",
        "std_behav": "ON 5s Popup",
        "std_par": "Key-ON",
        "rec_behav": "OFF",
        "display_msg": "MSG_088 - Top up washer fluid",
        "display_rec": "OFF",
    },
    {
        "ipc": "HIGH V10",
        "name": "RAIN SENSOR FAIL V1",
        "vfs": "VF150_V2_PJ3U",
        "hard_tt": "-",
        "soft_tt": "SOFT_RAIN_FAIL",
        "driven": "CAN",
        "std_behav": "ON 5s Popup",
        "std_par": "Fault transition",
        "rec_behav": "OFF",
        "display_msg": "MSG_089 - Auto wipers fault",
        "display_rec": "OFF",
    },
    # VF118
    {
        "ipc": "HIGH V10",
        "name": "SEAT BELT REMINDER FRONT V1",
        "vfs": "VF118_V3_PJ3U",
        "hard_tt": "TT_SBR_RED",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON / Blinking Stage 2",
        "std_par": "V > V_SBR_SPEED_TH",
        "rec_behav": "OFF",
        "display_msg": "MSG_018 - Fasten Seat Belt",
        "display_rec": "OFF",
    },
    {
        "ipc": "HIGH V10",
        "name": "SEAT BELT REMINDER REAR V1",
        "vfs": "VF118_V3_PJ3U",
        "hard_tt": "-",
        "soft_tt": "SOFT_REAR_SBR",
        "driven": "CAN",
        "std_behav": "30s Status Graphic",
        "std_par": "Buckle change",
        "rec_behav": "OFF",
        "display_msg": "MSG_019 - Rear Belt Unfastened",
        "display_rec": "OFF",
    },
    # VF210
    {
        "ipc": "HIGH V10",
        "name": "TPMS LOW PRESSURE V1",
        "vfs": "VF210_V1_PJ3U",
        "hard_tt": "TT_TPMS_AMBER",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady + Tire Graphic",
        "std_par": "Hyst P_TPMS_WARN_HYST",
        "rec_behav": "ON",
        "display_msg": "MSG_110 - Check Tire Pressure",
        "display_rec": "ON",
    },
    {
        "ipc": "HIGH V10",
        "name": "TPMS SYSTEM FAIL V1",
        "vfs": "VF210_V1_PJ3U",
        "hard_tt": "TT_TPMS_AMBER",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "Blink 75s then Steady ON",
        "std_par": "T_TPMS_BLINK_TIME",
        "rec_behav": "ON",
        "display_msg": "MSG_111 - TPMS Service Required",
        "display_rec": "ON",
    },
    # INTENCIONAL RAR #7: Indicação ÓRFÃ na planilha Indications (referencia VF199 que não existe no pacote IPC)
    {
        "ipc": "HIGH V10",
        "name": "GLOW PLUG PRE-HEATING V1",
        "vfs": "VF199_V1_P363",
        "hard_tt": "TT_GLOW_AMBER",
        "soft_tt": "-",
        "driven": "CAN",
        "std_behav": "ON Steady",
        "std_par": "Diesel Pre-Heat",
        "rec_behav": "OFF",
        "display_msg": "MSG_199 - Wait for pre-heating",
        "display_rec": "OFF",
    },
    # INTENCIONAL RAR #8: Indicação ÓRFÃ herdada de projeto antigo sem VF preenchida na coluna C
    {
        "ipc": "HIGH V10",
        "name": "HILL DESCENT CONTROL FAIL V1",
        "vfs": "",
        "hard_tt": "-",
        "soft_tt": "SOFT_HDC_AMBER",
        "driven": "CAN",
        "std_behav": "ON 5s Popup",
        "std_par": "Legacy carryover",
        "rec_behav": "OFF",
        "display_msg": "MSG_204 - HDC Unavailable",
        "display_rec": "OFF",
    },
]


# Linhas da planilha Section B - Software Requirements / IPC_Parameters_Table_PJ3U_R5.xlsx
# Colunas exibidas no vídeo (19:15):
# Parameter Name | First Trial Value | Range Value | Resolution | Unit | Parameter Description | Reference | 363 | 376 | J3U | Column Used by IPC HW
PARAMETERS_EXCEL_ROWS = [
    {
        "name": "tREC",
        "first_trial": "2,5",
        "range": "0,1÷5",
        "resolution": "0,1",
        "unit": "s",
        "description": "Time-out of the signals received for node recovery (tolerance ± 10%).",
        "reference": "Diagnosis / VF139",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "T_AIRBAG_STARTUP_CHK",
        "first_trial": "4,0",
        "range": "2,0÷6,0",
        "resolution": "0,5",
        "unit": "s",
        "description": "Duration of the Airbag telltale bulb check phase after Key-ON.",
        "reference": "VF139",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "V_TH_HB",
        "first_trial": "5",
        "range": "3÷15",
        "resolution": "1",
        "unit": "km/h",
        "description": "Vehicle speed threshold to trigger Hand Brake engaged with vehicle in movement warning.",
        "reference": "VF143",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "T_HB_WARN",
        "first_trial": "1,5",
        "range": "0,5÷5,0",
        "resolution": "0,5",
        "unit": "s",
        "description": "Debounce time with vehicle speed > V_TH_HB before acoustic warning activation.",
        "reference": "VF143",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    # Nota: T_HB_CHIME_REPEAT (VF143) e D_SBR_DIST_TH (VF118) foram omitidos intencionalmente para simular RAR de parâmetro EOL ausente na planilha!
    {
        "name": "T_LBF_FILTER",
        "first_trial": "2,0",
        "range": "0,5÷5,0",
        "resolution": "0,1",
        "unit": "s",
        "description": "Filtering time on CAN signal BrakeFluidLevelSts before turning ON red telltale.",
        "reference": "VF145",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "V_SBR_SPEED_TH",
        "first_trial": "20",
        "range": "10÷30",
        "resolution": "1",
        "unit": "km/h",
        "description": "Vehicle speed threshold for Seat Belt Reminder second stage visual and audible alarm.",
        "reference": "VF118",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "T_SBR_CHIME_DURATION",
        "first_trial": "90",
        "range": "30÷180",
        "resolution": "5",
        "unit": "s",
        "description": "Total duration of SBR acoustic chime sequence before automatic muting.",
        "reference": "VF118",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        # INTENCIONAL RAR #6: Parâmetro presente na planilha, mas com First Trial Value VAZIO ("") / divergente da VF210 (que define 75 s)
        "name": "T_TPMS_BLINK_TIME",
        "first_trial": "",
        "range": "60÷90",
        "resolution": "5",
        "unit": "s",
        "description": "Blinking duration of TPMS malfunction telltale before remaining continuously illuminated.",
        "reference": "VF210",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    {
        "name": "P_TPMS_WARN_HYST",
        "first_trial": "0,2",
        "range": "0,1÷0,5",
        "resolution": "0,05",
        "unit": "bar",
        "description": "Pressure hysteresis to clear low tire pressure warning on IPC display.",
        "reference": "VF210",
        "p363": "x",
        "p376": "x",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    # INTENCIONAL RAR #9: Parâmetro ÓRFÃO na planilha Parameters Table (herdado de outro projeto e não utilizado por nenhuma VF)
    {
        "name": "T_OIL_PRESS_DELAY_LEGACY",
        "first_trial": "3,0",
        "range": "1,0÷10,0",
        "resolution": "0,5",
        "unit": "s",
        "description": "Legacy oil pressure switch delay timer from P363 platform (not used in J3U VFs).",
        "reference": "Legacy P363",
        "p363": "x",
        "p376": "",
        "pj3u": "x",
        "hw_col": "ALL",
    },
    # INTENCIONAL RAR #10: Outro parâmetro ÓRFÃO na planilha Parameters Table
    {
        "name": "V_OVERSPEED_GCC_TH",
        "first_trial": "120",
        "range": "100÷140",
        "resolution": "5",
        "unit": "km/h",
        "description": "Fixed overspeed warning threshold for Middle East GCC market (no VF in LATAM package).",
        "reference": "VF177 (Not in package)",
        "p363": "",
        "p376": "",
        "pj3u": "x",
        "hw_col": "ALL",
    },
]


def set_cell_shading(cell, color_hex: str):
    shading_elm = OxmlElement("w:shd")
    shading_elm.set(qn("w:val"), "clear")
    shading_elm.set(qn("w:color"), "auto")
    shading_elm.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shading_elm)


def populate_vf_in_docx(doc: docx.Document, vf: dict):
    sec = vf["section_prefix"]
    h1 = doc.add_heading(f"{sec} {vf['title']} ({vf['vf_full_code']})", level=1)
    for run in h1.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    meta_p = doc.add_paragraph()
    meta_p.add_run("Document Code: ").bold = True
    meta_p.add_run(f"{vf['vf_full_code']}   |   ")
    meta_p.add_run("Author: ").bold = True
    meta_p.add_run(f"{vf['author']}   |   ")
    meta_p.add_run("Components Involved: ").bold = True
    meta_p.add_run(", ".join(vf["components_involved"]))

    doc.add_heading(f"{sec}.1 Purpose and Functional Overview", level=2)
    doc.add_paragraph(
        f"This Vehicle Function ({vf['vf_full_code']}) specifies the functional behavior, CAN signal processing, "
        f"visual/acoustic driver indications on the Instrument Panel Cluster (IPC), and end-of-line (EOL) calibration "
        f"parameters for {vf['title']}."
    )

    # Seção Indication Management (ex: 38.12 ou 62.12 conforme vídeo)
    doc.add_heading(f"{sec}.12 Indication Management", level=2)
    doc.add_paragraph(
        "The Instrument Panel Cluster (IPC) shall manage the following driver indications (telltales, pop-ups, and acoustic chimes) "
        "in accordance with the HMI specification (Section C - Indications):"
    )

    ind_table = doc.add_table(rows=1, cols=4)
    ind_table.style = "Table Grid"
    ind_headers = ["Indication Name", "Activation Condition", "Display Pop-up / Message", "Telltale Behavior"]
    hdr_cells = ind_table.rows[0].cells
    for i, text in enumerate(ind_headers):
        hdr_cells[i].text = text
        set_cell_shading(hdr_cells[i], "1F3864")
        for p in hdr_cells[i].paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(9)

    for ind in vf["indications"]:
        row_cells = ind_table.add_row().cells
        row_cells[0].text = ind["name"]
        row_cells[1].text = ind["condition"]
        row_cells[2].text = ind["Popup"]
        row_cells[3].text = ind["telltale"]
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)

    doc.add_paragraph()

    # Seção Configuration Parameters Table (ex: 38.14.1 conforme vídeo)
    doc.add_heading(f"{sec}.14 Configuration Parameters", level=2)
    doc.add_heading(f"{sec}.14.1 Configuration Parameters Table", level=3)
    doc.add_paragraph(
        "The table below defines the calibration and configuration parameters required by this Vehicle Function:"
    )

    param_headers = [
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
    param_table = doc.add_table(rows=1, cols=len(param_headers))
    param_table.style = "Table Grid"
    p_hdr_cells = param_table.rows[0].cells
    for i, text in enumerate(param_headers):
        p_hdr_cells[i].text = text
        set_cell_shading(p_hdr_cells[i], "2F5597")
        for p in p_hdr_cells[i].paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(8)

    if vf["parameters"]:
        for par in vf["parameters"]:
            r_cells = param_table.add_row().cells
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
            for i, val in enumerate(vals):
                r_cells[i].text = val
                for p in r_cells[i].paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8)
    else:
        # Linha vazia para VF sem parâmetros (como mostrado no vídeo aos 18:22)
        r_cells = param_table.add_row().cells
        for i in range(len(param_headers)):
            r_cells[i].text = "-"

    doc.add_paragraph()


def create_vf_pdf(pdf_path: Path, vfs_list: list[dict], title_header: str):
    """Gera versão em PDF estruturada das VFs usando PyMuPDF (fitz)."""
    pdf_doc = fitz.open()
    for vf in vfs_list:
        page = pdf_doc.new_page(width=595, height=842)  # A4
        sec = vf["section_prefix"]
        y = 45

        # Cabeçalho Stellantis
        page.insert_text(
            (40, y),
            f"STELLANTIS - EE SYSTEM REQUIREMENTS ({title_header})",
            fontsize=9,
            color=(0.3, 0.3, 0.3),
        )
        y += 22
        page.insert_text(
            (40, y),
            f"{sec} {vf['title']} ({vf['vf_full_code']})",
            fontsize=13,
            color=(0.12, 0.22, 0.39),
        )
        y += 18
        page.insert_text(
            (40, y),
            f"Document Code: {vf['vf_full_code']} | Components: {', '.join(vf['components_involved'])}",
            fontsize=9,
            color=(0.2, 0.2, 0.2),
        )
        y += 24

        # Seção Indication Management
        page.insert_text(
            (40, y),
            f"{sec}.12 Indication Management",
            fontsize=11,
            color=(0.12, 0.22, 0.39),
        )
        y += 16
        for ind in vf["indications"]:
            line = f"IND_ROW | {ind['name']} | {ind['condition']} | {ind['Popup']} | {ind['telltale']}"
            page.insert_text((45, y), line[:105], fontsize=8.5, color=(0, 0, 0))
            y += 14

        y += 14
        page.insert_text(
            (40, y),
            f"{sec}.14.1 Configuration Parameters Table",
            fontsize=11,
            color=(0.12, 0.22, 0.39),
        )
        y += 16
        page.insert_text(
            (45, y),
            "ID | Parameter Name | First Trial Value | Range Value | Resolution | Unit | Component | Implementation | XCP | Impact | Description",
            fontsize=7.5,
            color=(0.18, 0.33, 0.59),
        )
        y += 14

        if vf["parameters"]:
            for par in vf["parameters"]:
                pline = (
                    f"PAR_ROW | {par['id']} | {par['name']} | {par['first_trial']} | {par['range']} | "
                    f"{par['resolution']} | {par['unit']} | {par['component']} | {par['implementation']} | "
                    f"{par['xcp']} | {par['impact']} | {par['description']}"
                )
                page.insert_text((45, y), pline[:115], fontsize=7.5, color=(0, 0, 0))
                y += 13
        else:
            page.insert_text((45, y), "PAR_ROW | - | - | - | - | - | - | - | - | - | - | -", fontsize=7.5)

    pdf_doc.save(str(pdf_path))
    pdf_doc.close()


def generate_indications_excel(xlsx_path: Path):
    wb = openpyxl.Workbook()
    # Criar abas reais vistas no vídeo: Revisions, TFT 3.5 Inch. Monoc. LATAM, Full TFT 7 Inch. - LATAM, All IPC, Legenda
    ws_rev = wb.active
    ws_rev.title = "Revisions"
    ws_rev.append(["Rev", "Date", "Author", "Description"])
    ws_rev.append(["R1", "2026-02-10", "Stellantis HMI Team", "Initial release for J3U project"])
    ws_rev.append(["R4", "2026-10-28", "Stellantis HMI Team", "Updated telltale & popup mapping for PJ3U"])

    ws_mono = wb.create_sheet("TFT 3.5 Inch. Monoc. LATAM")
    ws_mono.append(["IPC", "Name", "VFs", "Notes"])
    ws_mono.append(["LOW V10", "Not applicable for HIGH V10 audit", "-", "Reference sheet"])

    ws_main = wb.create_sheet("Full TFT 7 Inch. - LATAM")
    ws_all = wb.create_sheet("All IPC")
    ws_leg = wb.create_sheet("Legenda")
    ws_leg.append(["Code", "Meaning"])
    ws_leg.append(["HIGH V10", "Full TFT 7 Inch Cluster LATAM"])
    ws_leg.append(["CAN", "Driven via CAN-FD / CAN High Speed network"])

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

    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for target_ws in (ws_main, ws_all):
        target_ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = target_ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for r_data in INDICATIONS_EXCEL_ROWS:
            row_vals = [
                r_data["ipc"],
                r_data["name"],
                r_data["vfs"],
                r_data["hard_tt"],
                r_data["soft_tt"],
                r_data["driven"],
                r_data["std_behav"],
                r_data["std_par"],
                r_data["rec_behav"],
                r_data["display_msg"],
                r_data["display_rec"],
            ]
            target_ws.append(row_vals)

        for row in target_ws.iter_rows(min_row=2, max_row=target_ws.max_row, max_col=len(headers)):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

        col_widths = [14, 38, 20, 18, 18, 10, 24, 20, 16, 30, 18]
        for idx, width in enumerate(col_widths, start=1):
            col_letter = openpyxl.utils.get_column_letter(idx)
            target_ws.column_dimensions[col_letter].width = width

    wb.save(str(xlsx_path))


def generate_parameters_excel(xlsx_path: Path):
    wb = openpyxl.Workbook()
    ws_rev = wb.active
    ws_rev.title = "Revisions"
    ws_rev.append(["Revision", "Date", "Author", "Change Log"])
    ws_rev.append(["R5", "2026-11-02", "Stellantis SW Spec Team", "Parameters Table update for PJ3U SOP"])

    ws_par = wb.create_sheet("Parameters Table")
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
    ws_par.append(headers)

    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws_par.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for p_data in PARAMETERS_EXCEL_ROWS:
        ws_par.append(
            [
                p_data["name"],
                p_data["first_trial"],
                p_data["range"],
                p_data["resolution"],
                p_data["unit"],
                p_data["description"],
                p_data["reference"],
                p_data["p363"],
                p_data["p376"],
                p_data["pj3u"],
                p_data["hw_col"],
            ]
        )

    for row in ws_par.iter_rows(min_row=2, max_row=ws_par.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    col_widths = [26, 18, 14, 12, 10, 56, 22, 8, 8, 8, 22]
    for idx, width in enumerate(col_widths, start=1):
        col_letter = openpyxl.utils.get_column_letter(idx)
        ws_par.column_dimensions[col_letter].width = width

    wb.save(str(xlsx_path))


def generate_all_mock_files() -> dict:
    VF_INDIVIDUAL_DIR.mkdir(parents=True, exist_ok=True)
    INDICATIONS_DIR.mkdir(parents=True, exist_ok=True)
    PARAMETERS_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Documento Word consolidado VFS_IPC_PJ3U_2A.docx e PDF consolidado
    consolidated_docx_path = VF_DIR / "VFS_IPC_PJ3U_2A.docx"
    consolidated_doc = docx.Document()
    title_p = consolidated_doc.add_heading("STELLANTIS - IPC VEHICLE FUNCTIONS SPECIFICATION PACKAGE (PJ3U)", 0)
    for r in title_p.runs:
        r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    consolidated_doc.add_paragraph(
        "Project: J3U - 5870 | Subsystem: IPC - Instrument Panel Cluster 10 (CAN) | Package Release: 2A"
    )

    # 2. Gerar também cada VF individual (.docx e .pdf) na subpasta IPC_J3U_2A/
    individual_files = []
    for vf in VFS_DATA:
        populate_vf_in_docx(consolidated_doc, vf)

        ind_doc = docx.Document()
        populate_vf_in_docx(ind_doc, vf)
        ind_docx_path = VF_INDIVIDUAL_DIR / f"{vf['filename_stem']}.docx"
        ind_doc.save(str(ind_docx_path))

        ind_pdf_path = VF_INDIVIDUAL_DIR / f"{vf['filename_stem']}.pdf"
        create_vf_pdf(ind_pdf_path, [vf], title_header=vf["vf_full_code"])
        individual_files.append(str(ind_docx_path))

    consolidated_doc.save(str(consolidated_docx_path))
    consolidated_pdf_path = VF_DIR / "VFS_IPC_PJ3U_2A.pdf"
    create_vf_pdf(consolidated_pdf_path, VFS_DATA, title_header="VFS_IPC_PJ3U_2A")

    # 3. Gerar Planilha de Indicações
    indications_xlsx_path = INDICATIONS_DIR / "Indications_J3U_R4.xlsx"
    generate_indications_excel(indications_xlsx_path)

    # 4. Gerar Planilha de Parâmetros
    parameters_xlsx_path = PARAMETERS_DIR / "IPC_Parameters_Table_PJ3U_R5.xlsx"
    generate_parameters_excel(parameters_xlsx_path)

    return {
        "project_root": str(MOCK_ROOT),
        "consolidated_docx": str(consolidated_docx_path),
        "consolidated_pdf": str(consolidated_pdf_path),
        "individual_vfs_dir": str(VF_INDIVIDUAL_DIR),
        "indications_xlsx": str(indications_xlsx_path),
        "parameters_xlsx": str(parameters_xlsx_path),
    }


if __name__ == "__main__":
    paths = generate_all_mock_files()
    print("Arquivos fictícios (Mock Data) gerados com sucesso:")
    for k, v in paths.items():
        print(f"  - {k}: {v}")
