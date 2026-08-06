import re
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


st.set_page_config(page_title="GCTS - Cruce Objetivos SAFA/SACA/SANA", layout="wide")


APP_TITLE = "Cruce automático: Lista de tráfico (NOP) + Objetivos SAFA/SACA/SANA"
APP_CAPTION = (
    "Sube el PDF de la lista de tráfico (NOP Eurocontrol, formato ARCID) y el Excel maestro "
    "de Objetivos SAFA/SACA/SANA/Matrículas. La app cruza cada vuelo por el prefijo ARCID "
    "y genera un Excel y un PDF enriquecidos."
)

ATYP_CODES = [
    "DA42", "A20N", "A21N", "A320", "A321", "A319", "AT76", "B738", "B38M", "C680",
    "A332", "A333", "A350", "A330", "A339", "T380", "A380", "A300", "A306", "B772",
    "B763", "B764", "B788", "B789", "B748", "B77L", "E295", "E550", "E190", "E170",
    "E175", "E145", "E135", "CRJ9", "CRJ7", "CRJ2", "CRJX", "BCS1", "BCS3", "SB20",
    "F900", "GLF6",
]

REG_PREFIXES = [
    "9H", "9M", "9V", "9A", "9K", "9G", "4X", "4R", "4L", "HB", "HA", "HS", "HL", "HK", "HP", "HZ",
    "LX", "LY", "LZ", "LV", "LN", "EC", "EI", "EK", "EP", "ES", "ET", "EW", "EY", "OE", "OO", "OY",
    "OH", "OK", "OM", "OB", "PH", "PK", "PP", "PR", "PT", "PZ", "SE", "SP", "ST", "SU", "SX", "TC",
    "TF", "TG", "TJ", "TN", "TR", "TS", "TU", "TY", "TZ", "UK", "UR", "VH", "VN", "VP", "VQ", "VT",
    "XA", "XB", "XC", "YI", "YJ", "YK", "YL", "YR", "YU", "YV", "ZA", "ZK", "ZP", "ZS", "CN", "CS",
    "CC", "CP", "CU", "CX", "G", "D", "F", "N", "B", "I", "J", "H", "P", "V", "Z", "C",
]
REG_PREFIXES_SORTED = sorted(set(REG_PREFIXES), key=len, reverse=True)
REG_EXPECTED_LEN = {"A7": 5, "A6": 5, "EC": 5}

ATYP_PAT = re.compile(r"(" + "|".join(ATYP_CODES) + r")")
TTV_PAT = re.compile(r"[A-Za-z]\s?\d{3}\s?\d{2}-")
TIME_PAT_F1 = re.compile(r"^(\d{2}:\d{2})A\s*(.*)$")
FLIGHT_LINE_PAT_F1 = re.compile(r"^\d{2}:\d{2}A")
FLIGHT_START_PAT_F2 = re.compile(r"(\d{2}:\d{2})([AEC])(?=\s?(?:[A-Z]{1,4}\s?)?[A-Z]{2,4}\d)")
EXPECTED_TOTAL_PAT = re.compile(r"-\s*(\d+)\s*Flights", re.IGNORECASE)

CURRENT_PDF_ICAOS: set = set()


@dataclass
class CrossResult:
    tipo: str
    operador: str
    inspecciones: Optional[object]
    objetivo: Optional[object]
    restantes: Optional[object]
    ultima: str
    fuente: str


def choose_best_registration(reg_raw: str) -> str:
    if not reg_raw:
        return ""
    candidates = []
    for prefix in REG_PREFIXES_SORTED:
        if reg_raw.startswith(prefix) and len(reg_raw) > len(prefix):
            suffix = reg_raw[len(prefix):]
            if 2 <= len(suffix) <= 4 and suffix.isalnum():
                score = 0
                if len(suffix) > 3:
                    score += 1
                expected = REG_EXPECTED_LEN.get(prefix)
                if expected and len(suffix) == expected - len(prefix):
                    score -= 1
                candidates.append((score, f"{prefix}-{suffix}"))
    if candidates:
        return sorted(candidates, key=lambda x: x[0])[0][1]
    return reg_raw


def looks_like_format2(text_sample: str) -> bool:
    for line in text_sample.splitlines():
        line = line.strip()
        if re.match(r"^\d{2}:\d{2}[AEC]", line):
            return True
        if re.match(r"^\d{2}:\d{2}A\s", line):
            return False
    return False


def extract_reg_airports(block: str) -> Tuple[str, str, str]:
    compact = block.replace(" ", "").replace(">", "")
    if not compact:
        return "", "", ""

    airport_hits = [
        (m.start(), m.group(0))
        for m in re.finditer(r"[A-Z]{4}", compact)
        if m.group(0) in CURRENT_PDF_ICAOS
    ]
    if len(airport_hits) >= 2:
        adep = airport_hits[-2][1]
        ades = airport_hits[-1][1]
        reg_raw = compact[:airport_hits[-2][0]]
        return reg_raw, adep, ades

    generic_hits = list(re.finditer(r"[A-Z]{4}", compact))
    if len(generic_hits) >= 2:
        adep = generic_hits[-2].group(0)
        ades = generic_hits[-1].group(0)
        reg_raw = compact[:generic_hits[-2].start()]
        return reg_raw, adep, ades

    if len(compact) >= 8:
        return compact[:-8], compact[-8:-4], compact[-4:]
    return "", compact[:4], compact[4:8]


def parse_format1(raw_lines: List[str]) -> pd.DataFrame:
    merged, current = [], None
    for line in raw_lines:
        if FLIGHT_LINE_PAT_F1.match(line):
            if current:
                merged.append(current)
            current = line
        elif current:
            current += " " + line
    if current:
        merged.append(current)

    rows = []
    for line in merged:
        match = TIME_PAT_F1.match(line)
        if not match:
            continue
        hora, rest = match.group(1), match.group(2)
        tokens = rest.split(" ")
        while tokens and re.fullmatch(r"[A-Z]{1,4}", tokens[0]) and not re.search(r"\d", tokens[0]):
            tokens.pop(0)
        rest = " ".join(tokens)
        atyp_match = ATYP_PAT.search(rest)

        if not atyp_match:
            first_token = rest.split(" ")[0] if rest else ""
            rows.append({"Hora": hora, "ARCID": first_token, "Aeronave": "", "Matricula": "", "ADEP": "", "ADES": "", "prefix3": first_token[:3]})
            continue

        arcid = rest[:atyp_match.start()].strip()
        atyp = atyp_match.group(1)
        remainder = rest[atyp_match.end():].replace(" ", "")
        matricula = choose_best_registration(remainder[:5]) if len(remainder) >= 5 else ""
        rows.append({
            "Hora": hora,
            "ARCID": arcid,
            "Aeronave": atyp,
            "Matricula": matricula,
            "ADEP": remainder[5:9],
            "ADES": remainder[9:13],
            "prefix3": re.match(r"^[A-Z]{3}", arcid).group(0) if re.match(r"^[A-Z]{3}", arcid) else arcid[:3],
        })
    return pd.DataFrame(rows)


def parse_one_flight_chunk(hora: str, rest: str) -> Optional[Dict[str, str]]:
    rest = rest.lstrip()
    atyp_match = ATYP_PAT.search(rest)
    if not atyp_match:
        return None

    prefix = rest[:atyp_match.start()]
    atyp = atyp_match.group(1)
    remainder = rest[atyp_match.end():]

    digit_match = re.search(r"\d", prefix)
    if digit_match:
        alpha_run = prefix[:digit_match.start()]
        digit_start = digit_match.start()
    else:
        alpha_run = prefix
        digit_start = len(prefix)

    airline_code = alpha_run[-3:] if len(alpha_run) >= 3 else alpha_run
    arcid = (airline_code + prefix[digit_start:]).strip()

    anchor = TTV_PAT.search(remainder)
    reg_airport_block = remainder[:anchor.start()] if anchor else remainder
    reg_raw, adep, ades = extract_reg_airports(reg_airport_block)

    return {
        "Hora": hora,
        "ARCID": arcid,
        "Aeronave": atyp,
        "Matricula": choose_best_registration(reg_raw),
        "ADEP": adep,
        "ADES": ades,
        "prefix3": airline_code.strip(),
    }


def parse_format2(raw_lines: List[str]) -> pd.DataFrame:
    rows = []
    for line in raw_lines:
        matches = list(FLIGHT_START_PAT_F2.finditer(line))
        for index, match in enumerate(matches):
            hora = match.group(1)
            chunk_start = match.end()
            chunk_end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
            parsed = parse_one_flight_chunk(hora, line[chunk_start:chunk_end])
            if parsed:
                rows.append(parsed)
    return pd.DataFrame(rows)


def parse_pdf_flights(pdf_bytes: bytes) -> pd.DataFrame:
    global CURRENT_PDF_ICAOS
    raw_lines: List[str] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines.extend([line.strip() for line in text.splitlines() if line.strip()])

    CURRENT_PDF_ICAOS = set(re.findall(r"\b[A-Z]{4}\b", "\n".join(raw_lines)))
    full_text = "\n".join(raw_lines)
    return parse_format2(raw_lines) if looks_like_format2(full_text) else parse_format1(raw_lines)


def extract_expected_total(pdf_bytes: bytes) -> Optional[int]:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            match = EXPECTED_TOTAL_PAT.search(text)
            if match:
                return int(match.group(1))
    return None


def extract_raw_arcid_candidates(pdf_bytes: bytes) -> pd.DataFrame:
    raw_lines: List[str] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines.extend([line.strip() for line in text.splitlines() if line.strip()])

    full_text = "\n".join(raw_lines)
    rows = []
    if looks_like_format2(full_text):
        for line in raw_lines:
            matches = list(FLIGHT_START_PAT_F2.finditer(line))
            for index, match in enumerate(matches):
                hora = match.group(1)
                chunk_start = match.end()
                chunk_end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
                chunk = line[chunk_start:chunk_end].lstrip()
                atyp_match = ATYP_PAT.search(chunk)
                prefix = chunk[:atyp_match.start()] if atyp_match else chunk[:12]
                digit_match = re.search(r"\d", prefix)
                alpha_run = prefix[:digit_match.start()] if digit_match else prefix
                digit_start = digit_match.start() if digit_match else len(prefix)
                airline_code = alpha_run[-3:] if len(alpha_run) >= 3 else alpha_run
                rows.append({
                    "Hora": hora,
                    "ARCID_guess": (airline_code + prefix[digit_start:]).strip(),
                    "parsed_ok": atyp_match is not None,
                })
    else:
        for line in raw_lines:
            match = TIME_PAT_F1.match(line)
            if not match:
                continue
            hora, rest = match.group(1), match.group(2)
            atyp_match = ATYP_PAT.search(rest)
            rows.append({
                "Hora": hora,
                "ARCID_guess": rest[:atyp_match.start()].strip() if atyp_match else (rest.split(" ")[0] if rest else ""),
                "parsed_ok": atyp_match is not None,
            })
    return pd.DataFrame(rows)


def load_sheet(xlsx_bytes: bytes, sheetname: str, max_col: int = 60):
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb[sheetname]
    headers, body = None, []
    for i, row in enumerate(ws.iter_rows(min_col=1, max_col=max_col, values_only=True)):
        if i == 0:
            headers = row
        else:
            body.append(row)
    wb.close()
    return headers, body


def fmt_date(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


def build_master_maps(xlsx_bytes: bytes):
    h_icao, r_icao = load_sheet(xlsx_bytes, "ICAO CODE", max_col=2)
    icao_map = {}
    for row in r_icao:
        name, code = row[0], row[1]
        if code is not None and name is not None:
            icao_map[str(code).strip().upper()] = str(name).strip()

    h1, r1 = load_sheet(xlsx_bytes, "Layer 1 Objectives")
    idx1 = {h: i for i, h in enumerate(h1) if h}
    l1_map = {}
    for row in r1:
        code = row[idx1.get("3LC")] if "3LC" in idx1 else None
        if code is None:
            continue
        code = str(code).strip().upper()
        l1_map[code] = {
            "operator": row[idx1.get("Operator Name")] if "Operator Name" in idx1 else "",
            "done": row[idx1.get("Progress")] if "Progress" in idx1 else None,
            "objective": row[idx1.get("Mean Target")] if "Mean Target" in idx1 else None,
            "remaining": row[idx1.get("Remaining")] if "Remaining" in idx1 else None,
            "last": row[idx1.get("Last inspection")] if "Last inspection" in idx1 else None,
        }

    h2, r2 = load_sheet(xlsx_bytes, "Layer 2 Objectives")
    idx2 = {h: i for i, h in enumerate(h2) if h}
    obj_key = next((k for k in idx2 if "OBJECTIVE" in str(k)), None)
    last_sp_key = next((k for k in idx2 if "LAST INSPECTION SPAIN" in str(k)), None)
    last_eu_key = next((k for k in idx2 if "LAST INSPECTION EUROPE" in str(k)), None)
    l2_map = {}
    for row in r2:
        code = row[idx2.get("3LC")] if "3LC" in idx2 else None
        if code is None:
            continue
        code = str(code).strip().upper()
        l2_map[code] = {
            "operator": row[idx2.get("OPERATOR L2")] if "OPERATOR L2" in idx2 else "",
            "done": row[idx2.get("DONE")] if "DONE" in idx2 else None,
            "objective": row[idx2.get(obj_key)] if obj_key else None,
            "remaining": row[idx2.get("REMAINING")] if "REMAINING" in idx2 else None,
            "last": (row[idx2.get(last_sp_key)] if last_sp_key else None) or (row[idx2.get(last_eu_key)] if last_eu_key else None),
        }

    hs, rs = load_sheet(xlsx_bytes, "SANA")
    idxs = {h: i for i, h in enumerate(hs) if h}
    sana_map = {}
    for row in rs:
        code = row[idxs.get("3LC")] if "3LC" in idxs else None
        if code is None:
            continue
        code = str(code).strip().upper()
        sana_map[code] = {
            "operator": row[idxs.get("Operator Name")] if "Operator Name" in idxs else "",
            "done": row[idxs.get("Inspections 2026")] if "Inspections 2026" in idxs else None,
            "objective": row[idxs.get("Objective 2026")] if "Objective 2026" in idxs else None,
            "remaining": row[idxs.get("Remaining inspections")] if "Remaining inspections" in idxs else None,
            "last": row[idxs.get("Last Inspection Date")] if "Last Inspection Date" in idxs else None,
        }

    hm, rm = load_sheet(xlsx_bytes, "Matrículas", max_col=10)
    idxm = {h: i for i, h in enumerate(hm) if h}
    matriculas_map = {}
    reg_key = next((k for k in idxm if str(k).strip().lower() in ["matricula", "matrícula", "registration"]), None)
    op_key = next((k for k in idxm if "operator" in str(k).strip().lower()), None)
    if reg_key and op_key:
        for row in rm:
            reg = row[idxm.get(reg_key)]
            op = row[idxm.get(op_key)]
            if reg and op:
                matriculas_map[str(reg).strip().upper()] = str(op).strip()

    return icao_map, l1_map, l2_map, sana_map, matriculas_map


def choose_effective_operator(prefix3: str, matricula: str, icao_map: Dict[str, str], matriculas_map: Dict[str, str]) -> Tuple[str, str]:
    matricula_norm = str(matricula).strip().upper() if matricula else ""
    if matricula_norm and matricula_norm in matriculas_map:
        return matriculas_map[matricula_norm], "Operador matrícula"
    prefix3_norm = str(prefix3).strip().upper() if prefix3 else ""
    if prefix3_norm in icao_map:
        return icao_map[prefix3_norm], "Operador prefijo ICAO"
    return "", "Sin operador"


def cross_reference(row: pd.Series, maps) -> CrossResult:
    icao_map, l1_map, l2_map, sana_map, matriculas_map = maps
    operador, fuente_operador = choose_effective_operator(row.get("prefix3", ""), row.get("Matricula", ""), icao_map, matriculas_map)
    prefix3 = str(row.get("prefix3", "")).strip().upper()

    if prefix3 in l1_map:
        item = l1_map[prefix3]
        return CrossResult("Layer 1", str(item.get("operator") or operador), item.get("done"), item.get("objective"), item.get("remaining"), fmt_date(item.get("last")), f"{fuente_operador} + Layer 1")
    if prefix3 in sana_map:
        item = sana_map[prefix3]
        return CrossResult("SANA", str(item.get("operator") or operador), item.get("done"), item.get("objective"), item.get("remaining"), fmt_date(item.get("last")), f"{fuente_operador} + SANA")
    if prefix3 in l2_map:
        item = l2_map[prefix3]
        return CrossResult("Layer 2", str(item.get("operator") or operador), item.get("done"), item.get("objective"), item.get("remaining"), fmt_date(item.get("last")), f"{fuente_operador} + Layer 2")
    return CrossResult("No encontrado", operador, None, None, None, "", fuente_operador)


def enrich_flights(df: pd.DataFrame, maps) -> pd.DataFrame:
    enriched = df.copy()
    results = enriched.apply(lambda row: cross_reference(row, maps), axis=1)
    enriched["Operador (maestro)"] = [r.operador for r in results]
    enriched["Tipo objetivo"] = [r.tipo for r in results]
    enriched["Inspecciones realizadas"] = [r.inspecciones for r in results]
    enriched["Objetivo 2026"] = [r.objetivo for r in results]
    enriched["Restantes"] = [r.restantes for r in results]
    enriched["Última inspección"] = [r.ultima for r in results]
    enriched["Fuente cruce"] = [r.fuente for r in results]
    return enriched


def build_excel(df: pd.DataFrame, fecha_str: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cruce"

    title_font = Font(bold=True, size=14, color="FFFFFF")
    header_font = Font(bold=True, color="FFFFFF")
    body_font = Font(size=10)
    border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    ws.merge_cells("B2:N2")
    ws["B2"] = f"Cruce tráfico NOP + Objetivos SAFA/SACA/SANA ({fecha_str})"
    ws["B2"].font = title_font
    ws["B2"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = list(df.columns)
    header_row = 4
    for col_num, header in enumerate(headers, start=2):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.font = header_font
        cell.fill = PatternFill("solid", fgColor="4F81BD")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    centered = {"Hora", "ARCID", "Aeronave", "Matricula", "ADEP", "ADES", "Tipo objetivo", "Inspecciones realizadas", "Objetivo 2026", "Restantes", "Última inspección"}
    for row_num, (_, row) in enumerate(df.iterrows(), start=header_row + 1):
        for col_num, header in enumerate(headers, start=2):
            value = None if pd.isna(row[header]) else row[header]
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center") if header in centered else Alignment(horizontal="left", vertical="center", indent=1)

    last_row = header_row + len(df)
    last_col = 1 + len(headers)
    last_col_letter = get_column_letter(last_col)
    widths = {"Hora": 8, "ARCID": 12, "Aeronave": 11, "Matricula": 12, "ADEP": 8, "ADES": 8, "Operador (maestro)": 42, "Tipo objetivo": 14, "Inspecciones realizadas": 12, "Objetivo 2026": 12, "Restantes": 10, "Última inspección": 16, "Fuente cruce": 24}
    for col_num, header in enumerate(headers, start=2):
        ws.column_dimensions[get_column_letter(col_num)].width = widths.get(header, 14)

    ws.freeze_panes = f"B{header_row + 1}"
    ws.auto_filter.ref = f"B{header_row}:{last_col_letter}{last_row}"

    tipo_col_letter = get_column_letter(2 + headers.index("Tipo objetivo"))
    rng = f"{tipo_col_letter}{header_row + 1}:{tipo_col_letter}{last_row}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Layer 1"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"SANA"'], fill=PatternFill("solid", fgColor="BDD7EE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Layer 2"'], fill=PatternFill("solid", fgColor="FFE699")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"No encontrado"'], fill=PatternFill("solid", fgColor="FFC7CE")))

    note_row = last_row + 3
    ws.cell(row=note_row, column=2, value="Fuente: PDF NOP Eurocontrol + Excel maestro Objetivos_SAFA_SACA_SANA_Matriculas.").font = Font(size=8, italic=True, color="808080")
    ws.cell(row=note_row + 1, column=2, value=f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}").font = Font(size=8, italic=True, color="808080")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_pdf(df: pd.DataFrame, fecha_str: str) -> BytesIO:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), leftMargin=10 * mm, rightMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
    styles = getSampleStyleSheet()
    title_style = styles["Heading1"]
    title_style.alignment = 1
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=8, leading=10)
    story = [Paragraph(f"Cruce tráfico NOP + Objetivos SAFA/SACA/SANA ({fecha_str})", title_style), Spacer(1, 6 * mm)]

    headers = list(df.columns)
    rows_per_page = 30
    for start in range(0, len(df), rows_per_page):
        chunk = df.iloc[start:start + rows_per_page]
        table_data = [[Paragraph(str(h), small) for h in headers]]
        for _, row in chunk.iterrows():
            table_data.append([Paragraph("" if pd.isna(row[h]) else str(row[h]), small) for h in headers])
        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(table)
        if start + rows_per_page < len(df):
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer


def apply_filters(result_df: pd.DataFrame) -> pd.DataFrame:
    st.markdown("### Filtros antes de descargar")
    col1, col2, col3 = st.columns(3)
    with col1:
        texto_busqueda = st.text_input("ARCID (texto libre)", "")
    with col2:
        ades_disponibles = sorted([a for a in result_df["ADES"].dropna().unique().tolist() if a])
        incluir_vacios = result_df["ADES"].isna().any() or (result_df["ADES"] == "").any()
        opciones_ades = (["(vacío)"] if incluir_vacios else []) + ades_disponibles
        ades_sel = st.multiselect("ADES (destino)", opciones_ades, default=[])
    with col3:
        operadores_disponibles = sorted([o for o in result_df["Operador (maestro)"].dropna().unique().tolist() if o])
        operadores_sel = st.multiselect("Operador", operadores_disponibles, default=[])

    col4, col5, col6 = st.columns(3)
    with col4:
        tipos_disponibles = sorted(result_df["Tipo objetivo"].dropna().unique().tolist())
        tipos_sel = st.multiselect("Tipo objetivo", tipos_disponibles, default=tipos_disponibles)
    with col5:
        restantes_num = pd.to_numeric(result_df["Restantes"], errors="coerce")
        max_restantes = int(restantes_num.max()) if restantes_num.notna().any() else 0
        restantes_range = st.slider("Restantes (rango)", 0, max(max_restantes, 1), (0, max(max_restantes, 1)))
    with col6:
        fechas_validas = pd.to_datetime(result_df["Última inspección"], errors="coerce")
        min_fecha = fechas_validas.min().date() if fechas_validas.notna().any() else datetime.now().date()
        max_fecha = fechas_validas.max().date() if fechas_validas.notna().any() else datetime.now().date()
        preset_sel = st.selectbox("Última inspección", ["Todas las fechas", "Última semana", "Último mes", "No en la última semana", "No en el último mes", "Rango personalizado"], index=0)
        fecha_range = st.date_input("Rango personalizado de fechas", value=(min_fecha, max_fecha)) if preset_sel == "Rango personalizado" else None

    filtered_df = result_df.copy()
    if texto_busqueda.strip():
        filtered_df = filtered_df[filtered_df["ARCID"].str.contains(texto_busqueda.strip(), case=False, na=False)]
    if ades_sel:
        quiere_vacios = "(vacío)" in ades_sel
        valores_reales = [a for a in ades_sel if a != "(vacío)"]
        mask_ades = filtered_df["ADES"].isin(valores_reales)
        if quiere_vacios:
            mask_ades = mask_ades | filtered_df["ADES"].isna() | (filtered_df["ADES"] == "")
        filtered_df = filtered_df[mask_ades]
    if operadores_sel:
        filtered_df = filtered_df[filtered_df["Operador (maestro)"].isin(operadores_sel)]
    if tipos_sel:
        filtered_df = filtered_df[filtered_df["Tipo objetivo"].isin(tipos_sel)]

    rest_num_full = pd.to_numeric(filtered_df["Restantes"], errors="coerce")
    filtered_df = filtered_df[rest_num_full.isna() | rest_num_full.between(restantes_range[0], restantes_range[1])]

    hoy = datetime.now().date()
    fechas_filtro = pd.to_datetime(filtered_df["Última inspección"], errors="coerce")
    if preset_sel == "Última semana":
        limite = hoy - pd.Timedelta(days=7)
        filtered_df = filtered_df[fechas_filtro.notna() & (fechas_filtro.dt.date >= limite) & (fechas_filtro.dt.date <= hoy)]
    elif preset_sel == "Último mes":
        limite = hoy - pd.Timedelta(days=30)
        filtered_df = filtered_df[fechas_filtro.notna() & (fechas_filtro.dt.date >= limite) & (fechas_filtro.dt.date <= hoy)]
    elif preset_sel == "No en la última semana":
        limite = hoy - pd.Timedelta(days=7)
        filtered_df = filtered_df[fechas_filtro.isna() | (fechas_filtro.dt.date < limite)]
    elif preset_sel == "No en el último mes":
        limite = hoy - pd.Timedelta(days=30)
        filtered_df = filtered_df[fechas_filtro.isna() | (fechas_filtro.dt.date < limite)]
    elif preset_sel == "Rango personalizado" and isinstance(fecha_range, tuple) and len(fecha_range) == 2:
        filtered_df = filtered_df[fechas_filtro.isna() | ((fechas_filtro.dt.date >= fecha_range[0]) & (fechas_filtro.dt.date <= fecha_range[1]))]
    return filtered_df


def render_app():
    st.title(APP_TITLE)
    st.caption(APP_CAPTION)

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. PDF de tráfico (NOP / ARCID)", type=["pdf"])
    with col2:
        xlsx_file = st.file_uploader("2. Excel maestro (Objetivos SAFA/SACA/SANA)", type=["xlsx"])

    run = st.button("Generar cruce", type="primary", disabled=not (pdf_file and xlsx_file))
    if not run:
        st.info("Sube ambos archivos y pulsa 'Generar cruce' para empezar.")
        return

    pdf_bytes = pdf_file.read()
    xlsx_bytes = xlsx_file.read()
    flights_df = parse_pdf_flights(pdf_bytes)
    maps = build_master_maps(xlsx_bytes)
    result_df = enrich_flights(flights_df, maps)
    raw_candidates_df = extract_raw_arcid_candidates(pdf_bytes)
    expected_total = extract_expected_total(pdf_bytes)
    fecha_str = datetime.now().strftime("%Y%m%d")

    st.success(f"Cruce completado con {len(result_df)} vuelos procesados.")
    if expected_total is not None:
        missing = max(expected_total - len(result_df), 0)
        coverage = (len(result_df) / expected_total * 100) if expected_total else 0
        st.caption(f"Cobertura: {len(result_df)} / {expected_total} vuelos ({coverage:.1f}%). Faltantes estimados: {missing}.")
        if missing > 0 and not raw_candidates_df.empty and st.button("Ver vuelos no detectados"):
            detected_arcids = set(result_df["ARCID"].astype(str).str.upper())
            raw_candidates_df["ARCID_norm"] = raw_candidates_df["ARCID_guess"].astype(str).str.upper()
            no_detectados = raw_candidates_df[~raw_candidates_df["ARCID_norm"].isin(detected_arcids)]
            if no_detectados.empty:
                st.success("No se han encontrado vuelos adicionales sin detectar.")
            else:
                st.dataframe(no_detectados[["Hora", "ARCID_guess", "parsed_ok"]].rename(columns={"ARCID_guess": "ARCID (detectado en texto crudo)", "parsed_ok": "Se pudo parsear tipo/aeropuertos"}), use_container_width=True)

    counts = result_df["Tipo objetivo"].value_counts()
    cols = st.columns(len(counts) if len(counts) > 0 else 1)
    for col, (tipo, n) in zip(cols, counts.items()):
        col.metric(tipo, n)

    filtered_df = apply_filters(result_df)
    st.caption(f"Mostrando {len(filtered_df)} de {len(result_df)} vuelos tras aplicar filtros.")
    st.dataframe(filtered_df, use_container_width=True, height=500)

    excel_buf_filtered = build_excel(filtered_df.reset_index(drop=True), fecha_str)
    pdf_buf_filtered = build_pdf(filtered_df.reset_index(drop=True), fecha_str)
    excel_buf_all = build_excel(result_df.reset_index(drop=True), fecha_str)
    pdf_buf_all = build_pdf(result_df.reset_index(drop=True), fecha_str)

    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button("Descargar Excel (con filtros aplicados)", data=excel_buf_filtered, file_name=f"GCTS_{fecha_str}_Enriquecido_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_b:
        st.download_button("Descargar PDF (con filtros aplicados)", data=pdf_buf_filtered, file_name=f"GCTS_{fecha_str}_Reconstruido_filtrado.pdf", mime="application/pdf")

    st.markdown("---")
    st.caption("¿Necesitas todo sin filtrar? Descárgalo aquí:")
    col_c, col_d = st.columns(2)
    with col_c:
        st.download_button("Descargar Excel completo (sin filtros)", data=excel_buf_all, file_name=f"GCTS_{fecha_str}_Enriquecido_completo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_d:
        st.download_button("Descargar PDF completo (sin filtros)", data=pdf_buf_all, file_name=f"GCTS_{fecha_str}_Reconstruido_completo.pdf", mime="application/pdf")


if __name__ == "__main__":
    render_app()
