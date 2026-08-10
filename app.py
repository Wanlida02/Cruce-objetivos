import re
import unicodedata
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

# ============================================================
# AIRCRAFT TYPE DESIGNATORS (ICAO Doc 8643) — comprehensive list
# ============================================================
# All ICAO type designators are exactly 2-4 alphanumeric characters. This list
# covers commercial, regional, business-jet and general-aviation types commonly
# seen in European IFR traffic (Eurocontrol NM Flight Lists). Verified against
# real NOP PDFs for LEMD and LEBL (591 flights, 0 unrecognized types after
# adding B77L, FA8X, GL7T, which were the only gaps found during testing).
ATYP_CODES_RAW = """
A124 A140 A148 A158 A19N A20N A21N A225 A306 A30B A310 A318 A319 A320 A321 A332 A333
A337 A338 A339 A342 A343 A345 A346 A359 A35K A388 A3ST A400 A748
AC90 AJ27 AN12 AN24 AN26 AN28 AN30 AN32 AN72 AT43 AT44 AT45 AT46 AT72 AT73 AT75 AT76 ATP
B190 B37M B38M B39M B3XM B461 B462 B463 B703 B712 B720 B721 B722 B732 B733 B734 B735
B736 B737 B738 B739 B741 B742 B743 B744 B748 B74R B74S B752 B753 B762 B763 B764 B772
B773 B778 B779 B77L B77W B788 B789 B78X BA11 BA46 BCS1 BCS3 BE20 BE30 BE33 BE35 BE36
BE40 BE58 BE76 BE95 BE99 BELF BER2 BLCF
C130 C17 C172 C177 C182 C206 C208 C210 C212 C25A C25B C25C C25M C310 C340 C408 C414
C421 C500 C510 C525 C550 C560 C56X C5M C650 C680 C68A C700 C750 C919 CL2T CL30 CL35
CL60 CN35 CRJ1 CRJ2 CRJ7 CRJ9 CRJX CVLT
D228 D328 DA20 DA40 DA42 DA62 DC10 DC85 DC86 DC87 DC91 DC92 DC93 DC94 DC95 DH8A DH8B
DH8C DH8D DHC5 DHC6 DHC7 DHT
E110 E120 E135 E145 E170 E175 E190 E195 E290 E295 E35L E50P E545 E550 E55P E75L E75S
EA50
F100 F27 F28 F2TH F406 F50 F70 F900 FA50 FA6X FA7X FA8X
G150 G159 G200 G280 G650 G73T GA5C GA6C GA7C GA8 GALX GL5T GL7T GLEX GLF4 GLF5 GLF6
H25B H25C HDJT
I114 IL18 IL62 IL76 IL86 IL96
J328 JS31 JS32 JS41
K35R
L101 L188 L410 LJ31 LJ35 LJ40 LJ45 LJ60 LJ70 LJ75
M20P M20T MD11 MD81 MD82 MD83 MD87 MD88 MD90 MU2
N262 NOMA
P180 P206 P208 P210 P28A P28B P28R P28T P68C P8 P92 PA28 PA34 PA44 PA46 PAY2 PC12 PC24
PRM1
RJ1H RJ70 RJ85
S601 SB20 SC7 SF34 SH33 SH36 SR20 SR22 SU95 SW4
T134 T154 T204 TBM7 TBM8 TBM9 TU34 TU54
WW24
Y12 YK40 YK42 YS11
""".split()
ATYP_CODES = sorted(set(ATYP_CODES_RAW), key=len, reverse=True)
ATYP_PAT = re.compile(r"(" + "|".join(re.escape(c) for c in ATYP_CODES) + r")")

# ============================================================
# AIRCRAFT NATIONALITY / REGISTRATION PREFIXES (ICAO Annex 7)
# ============================================================
# Sorted longest-first so multi-char prefixes (e.g. '9H', 'A6') are tried before
# ambiguous single-char ones. The single most important special case is the USA
# 'N'-number, which is NEVER hyphenated (e.g. N804AN stays N804AN), unlike every
# other country's registration, which gets a hyphen inserted right after the
# nationality prefix (e.g. 'ECLUB' -> 'EC-LUB', 'B324X' -> 'B-324X' for China).
REG_PREFIXES_ALL = [
    "A9C", "4YB", "9XR",
    "9A", "9G", "9H", "9J", "9K", "9L", "9M", "9N", "9Q", "9U", "9V", "9Y",
    "4K", "4L", "4O", "4R", "4X",
    "5A", "5B", "5H", "5N", "5R", "5T", "5U", "5V", "5W", "5X", "5Y",
    "6O", "6V", "6W", "6Y",
    "7O", "7P", "7Q", "7T",
    "8P", "8Q", "8R",
    "A2", "A3", "A4", "A5", "A6", "A7", "A8",
    "AP",
    "C2", "C5", "C6", "C9", "CC", "CN", "CP", "CR", "CS", "CU",
    "D2", "D4", "D6",
    "E7", "EC", "EI", "EJ", "EK", "EL", "EP", "ES", "ET", "EW", "EX", "EY", "EZ",
    "H4", "HA", "HB", "HC", "HH", "HI", "HK", "HL", "HP", "HR", "HS", "HZ",
    "JA", "JU", "JY",
    "LN", "LQ", "LR", "LV", "LX", "LY", "LZ",
    "MI", "MT",
    "OB", "OD", "OE", "OH", "OK", "OM", "OO", "OY",
    "PH", "PJ", "PK", "PP", "PR", "PT", "PU", "PZ",
    "RA", "RP",
    "S2", "S5", "S7", "S9", "SE", "SP", "ST", "SU", "SX",
    "T2", "T3", "T7", "T8", "T9", "TC", "TF", "TG", "TI", "TJ", "TL", "TN", "TR",
    "TS", "TT", "TU", "TY", "TZ",
    "UK", "UN", "UP", "UR",
    "V2", "V3", "V4", "V5", "V6", "V7", "V8", "VH", "VN", "VP", "VQ", "VR", "VT",
    "XA", "XB", "XC", "XT", "XU", "XY",
    "YA", "YI", "YJ", "YK", "YL", "YR", "YU", "YV",
    "Z3", "ZA", "ZJ", "ZK", "ZM", "ZP", "ZQ", "ZS",
    "B", "C", "D", "F", "G", "I", "J", "M", "N", "P", "T", "V", "Z",
]
REG_PREFIXES_SORTED = sorted(set(REG_PREFIXES_ALL), key=len, reverse=True)

TTV_PAT = re.compile(r"[A-Za-z]\s?\d{3}\s?\d{2}-")
TIME_PAT_F1 = re.compile(r"^(\d{2}:\d{2})A\s*(.*)$")
FLIGHT_LINE_PAT_F1 = re.compile(r"^\d{2}:\d{2}A")
# Broadened lookahead (vs. earlier versions) also matches single-letter ARCIDs
# formed from a US N-number used as callsign (e.g. business jets like N922DN),
# which the previous [A-Z]{2,4}\d-only lookahead was silently dropping.
FLIGHT_START_PAT_F2 = re.compile(r"(\d{2}:\d{2})([AEC])(?=\s?(?:[A-Z]{1,4}\s?)?[A-Z]{1,4}\d)")
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


def _norm_header(value) -> str:
    """Normalizes a column header for robust matching: strips accents,
    uppercases, collapses whitespace. This lets the app tolerate real-world
    header variations like 'OACI' vs '3LC', 'OPERADOR' vs 'Operator Name'."""
    s = str(value or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def _find_col(headers_norm: Dict[str, int], candidates: List[str]) -> Optional[int]:
    for cand in candidates:
        cand_norm = _norm_header(cand)
        for h_norm, idx in headers_norm.items():
            if cand_norm in h_norm:
                return idx
    return None


def hyphenate_registration(reg_raw: str) -> str:
    """Inserts a hyphen at the correct position based on real ICAO nationality
    registration-mark rules (ICAO Annex 7):
      - USA ('N'): NEVER hyphenated (e.g. N804AN, N922DN).
      - Every other country: 1-3 letter/digit nationality prefix + hyphen +
        suffix (e.g. 'ECLUB' -> 'EC-LUB', 'B324X' -> 'B-324X' for China,
        '9HALU' -> '9H-ALU' for Malta).
    Falls back to the raw string if no prefix matches (rare state/military
    aircraft with purely numeric identifiers)."""
    if not reg_raw:
        return ""
    reg_raw = reg_raw.strip()
    if re.match(r"^N\d", reg_raw):
        return reg_raw
    for prefix in REG_PREFIXES_SORTED:
        if reg_raw.startswith(prefix) and len(reg_raw) > len(prefix):
            return f"{prefix}-{reg_raw[len(prefix):]}"
    return reg_raw


def choose_best_registration(reg_raw: str) -> str:
    """Kept for backward compatibility with Format-1 traffic lists; delegates
    to the ICAO-rule-based hyphenation used everywhere else."""
    return hyphenate_registration(reg_raw)


def looks_like_format2(text_sample: str) -> bool:
    for line in text_sample.splitlines():
        line = line.strip()
        if re.match(r"^\d{2}:\d{2}[AEC]", line):
            return True
        if re.match(r"^\d{2}:\d{2}A\s", line):
            return False
    return False


def extract_reg_airports(block: str) -> Tuple[str, str, str]:
    """Extracts (registration, ADEP, ADES) from the fused text block that sits
    between the aircraft-type code and the Traffic-Volume anchor.

    KEY INSIGHT (verified against real Eurocontrol NM Flight List PDFs): ADEP
    and ADES are ALWAYS exactly 4 characters each, and they are always the
    LAST 8 characters of this block — this holds regardless of how pdfplumber
    fuses spacing, and regardless of whether the airport codes are ones we've
    seen before. Everything before those last 8 characters is the aircraft
    registration. This structural, position-based extraction is far more
    robust than trying to "recognize" ADEP/ADES from a whitelist of known ICAO
    airport codes, which would silently break on any airport not seen before."""
    compact = block.replace(" ", "").replace(">", "")
    if len(compact) >= 8:
        return compact[:-8], compact[-8:-4], compact[-4:]
    if len(compact) > 4:
        return "", compact[:4], compact[4:]
    return "", compact, ""


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
        matricula = hyphenate_registration(remainder[:5]) if len(remainder) >= 5 else ""
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
    """Parses a single flight's text chunk (already isolated from any
    neighbouring flights that pdfplumber may have fused onto the same
    physical line), starting right after the HH:MM[A|E|C] status indicator.

    Example input: 'LU RSC5SK AT76 EC-MIF GCTS GCLP A090 01-06:45+10:45 ...'

    Field layout (per Eurocontrol NM Flight List spec, confirmed against real
    PDFs): ARCID, ATYP(4 chars, from ICAO Doc 8643), REG, ADEP(4), ADES(4),
    then the Traffic-Volume anchor (e.g. 'A090 01-'). ATYP is found via the
    comprehensive dictionary above; REG/ADEP/ADES are then split by FIXED
    WIDTH (last 8 chars = ADEP+ADES) rather than by trying to recognize
    airport codes, which is what made earlier versions unreliable."""
    rest = rest.lstrip()

    atyp_match = ATYP_PAT.search(rest)
    if not atyp_match:
        return None

    prefix = rest[:atyp_match.start()]
    atyp = atyp_match.group(1)
    remainder = rest[atyp_match.end():]

    dm = re.search(r"\d", prefix)
    if dm:
        alpha_run = prefix[:dm.start()]
        digit_start = dm.start()
    else:
        alpha_run = prefix
        digit_start = len(prefix)

    airline_code = alpha_run[-3:] if len(alpha_run) >= 3 else alpha_run
    arcid = (airline_code + prefix[digit_start:]).strip()

    anchor = TTV_PAT.search(remainder)
    reg_airport_block = remainder[:anchor.start()] if anchor else remainder

    reg_raw, adep, ades = extract_reg_airports(reg_airport_block)
    reg = hyphenate_registration(reg_raw) if reg_raw else ""

    return {
        "Hora": hora, "ARCID": arcid.strip(), "Aeronave": atyp, "Matricula": reg,
        "ADEP": adep, "ADES": ades,
        "prefix3": airline_code.strip(),
    }


def parse_format2(raw_lines: List[str]) -> pd.DataFrame:
    """Parses NM/CFMU-style detailed traffic lists. Uses finditer over each raw
    text line to find EVERY 'HH:MM[A|E|C]' occurrence, because pdfplumber
    sometimes extracts two or three flight rows from the source PDF table as a
    single fused text line. Splitting on every match (instead of anchoring
    only at the start of the line) ensures no flight is silently dropped when
    this fusion happens."""
    rows = []
    for ln in raw_lines:
        matches = list(FLIGHT_START_PAT_F2.finditer(ln))
        if not matches:
            continue
        for i, m in enumerate(matches):
            hora = m.group(1)
            chunk_start = m.end()
            chunk_end = matches[i + 1].start() if i + 1 < len(matches) else len(ln)
            chunk = ln[chunk_start:chunk_end]
            parsed = parse_one_flight_chunk(hora, chunk)
            if parsed:
                rows.append(parsed)
    return pd.DataFrame(rows)


def extract_expected_total(pdf_bytes: bytes) -> Optional[int]:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            m = EXPECTED_TOTAL_PAT.search(txt)
            if m:
                return int(m.group(1))
    return None


def extract_raw_arcid_candidates(pdf_bytes: bytes) -> pd.DataFrame:
    """Scans the raw PDF text for every 'HH:MM[A|E|C]' flight-start marker,
    independent of the actual parsed DataFrame, to detect flights present in
    the source text but missing from the final parsed result (diagnostic
    signal for the "vuelos no detectados" button)."""
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        raw_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            raw_lines.extend([ln.strip() for ln in txt.splitlines() if ln.strip()])

    full_text = "\n".join(raw_lines)
    candidates = []
    if looks_like_format2(full_text):
        for ln in raw_lines:
            matches = list(FLIGHT_START_PAT_F2.finditer(ln))
            for i, m in enumerate(matches):
                hora = m.group(1)
                chunk_start = m.end()
                chunk_end = matches[i + 1].start() if i + 1 < len(matches) else len(ln)
                chunk = ln[chunk_start:chunk_end].lstrip()
                am = ATYP_PAT.search(chunk)
                prefix = chunk[:am.start()] if am else chunk[:12]
                dm = re.search(r"\d", prefix)
                alpha_run = prefix[:dm.start()] if dm else prefix
                digit_start = dm.start() if dm else len(prefix)
                airline_code = alpha_run[-3:] if len(alpha_run) >= 3 else alpha_run
                arcid_guess = airline_code + prefix[digit_start:]
                candidates.append({"Hora": hora, "ARCID_guess": arcid_guess.strip(), "parsed_ok": am is not None})
    else:
        merged = []
        current = None
        for ln in raw_lines:
            if FLIGHT_LINE_PAT_F1.match(ln):
                if current:
                    merged.append(current)
                current = ln
            elif current:
                current += " " + ln
        if current:
            merged.append(current)
        for ln in merged:
            m = TIME_PAT_F1.match(ln)
            if not m:
                continue
            hora, rest = m.group(1), m.group(2)
            tokens = rest.split(" ")
            while tokens and re.fullmatch(r"[A-Z]{1,4}", tokens[0]) and not re.search(r"\d", tokens[0]):
                tokens.pop(0)
            rest2 = " ".join(tokens)
            am = ATYP_PAT.search(rest2)
            arcid_guess = rest2[:am.start()].strip() if am else (rest2.split(" ")[0] if rest2 else "")
            candidates.append({"Hora": hora, "ARCID_guess": arcid_guess, "parsed_ok": am is not None})
    return pd.DataFrame(candidates)


def parse_pdf_flights(pdf_bytes: bytes) -> pd.DataFrame:
    """Extracts (Hora, ARCID, Aeronave, Matricula, ADEP, ADES) from either of
    the two known NOP/CFMU PDF traffic-list layouts, auto-detecting which one
    applies."""
    global CURRENT_PDF_ICAOS
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        raw_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            raw_lines.extend([ln.strip() for ln in txt.splitlines() if ln.strip()])

    full_text = "\n".join(raw_lines)
    CURRENT_PDF_ICAOS = set(re.findall(r"[A-Z]{4}", full_text))

    if looks_like_format2(full_text):
        df = parse_format2(raw_lines)
    else:
        df = parse_format1(raw_lines)

    return df


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


def find_sheet_name(available_sheets: List[str], keywords: List[str]) -> Optional[str]:
    """Finds the real worksheet name matching ALL keywords (accent- and
    case-insensitive), so small naming variations in the master Excel (e.g.
    'SANA' vs 'SANA Objectives') don't crash the app with a hard KeyError."""
    normalized = {s: _norm_header(s) for s in available_sheets}
    for sheet, upper in normalized.items():
        if all(_norm_header(kw) in upper for kw in keywords):
            return sheet
    return None


def fmt_date(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v)


def build_master_maps(xlsx_bytes: bytes):
    wb_names = openpyxl.load_workbook(BytesIO(xlsx_bytes), read_only=True, data_only=True).sheetnames

    icao_sheet = find_sheet_name(wb_names, ["ICAO"]) or find_sheet_name(wb_names, ["CODE"])
    l1_sheet = find_sheet_name(wb_names, ["LAYER 1"])
    l2_sheet = find_sheet_name(wb_names, ["LAYER 2"])
    sana_sheet = find_sheet_name(wb_names, ["SANA"])
    matriculas_sheet = find_sheet_name(wb_names, ["MATR"]) or find_sheet_name(wb_names, ["MATRICULAS"])

    missing = [
        label for label, sheet in [
            ("ICAO CODE", icao_sheet), ("Layer 1 Objectives", l1_sheet),
            ("Layer 2 Objectives", l2_sheet), ("SANA Objectives", sana_sheet),
            ("Matrículas", matriculas_sheet),
        ] if sheet is None
    ]
    if missing:
        st.warning(
            "No se encontraron en el Excel maestro las siguientes hojas esperadas: "
            + ", ".join(missing)
            + ". Hojas disponibles: " + ", ".join(wb_names)
            + ". Esas fuentes se omitirán del cruce, pero la app seguirá funcionando."
        )

    icao_map: Dict[str, str] = {}
    if icao_sheet:
        h_icao, r_icao = load_sheet(xlsx_bytes, icao_sheet, max_col=2)
        for row in r_icao:
            name, code = row[0], row[1]
            if code is not None and name is not None:
                icao_map[str(code).strip().upper()] = str(name).strip()

    l1_map: Dict[str, dict] = {}
    if l1_sheet:
        h1, r1 = load_sheet(xlsx_bytes, l1_sheet)
        h1_norm = {_norm_header(h): i for i, h in enumerate(h1) if h}
        code_idx = _find_col(h1_norm, ["3LC", "OACI"])
        op_idx = _find_col(h1_norm, ["OPERATOR NAME", "OPERADOR"])
        done_idx = _find_col(h1_norm, ["PROGRESS"])
        obj_idx = _find_col(h1_norm, ["MEAN TARGET"])
        rem_idx = _find_col(h1_norm, ["REMAINING"])
        last_idx = _find_col(h1_norm, ["LAST INSPECTION"])
        for row in r1:
            code = row[code_idx] if code_idx is not None else None
            if code is None:
                continue
            code = str(code).strip().upper()
            l1_map[code] = {
                "operator": row[op_idx] if op_idx is not None else "",
                "done": row[done_idx] if done_idx is not None else None,
                "objective": row[obj_idx] if obj_idx is not None else None,
                "remaining": row[rem_idx] if rem_idx is not None else None,
                "last": row[last_idx] if last_idx is not None else None,
            }

    l2_map: Dict[str, dict] = {}
    if l2_sheet:
        h2, r2 = load_sheet(xlsx_bytes, l2_sheet)
        h2_norm = {_norm_header(h): i for i, h in enumerate(h2) if h}
        code_idx = _find_col(h2_norm, ["3LC", "OACI"])
        op_idx = _find_col(h2_norm, ["OPERATOR L2", "OPERADOR"])
        done_idx = _find_col(h2_norm, ["DONE"])
        obj_idx = _find_col(h2_norm, ["OBJECTIVE"])
        rem_idx = _find_col(h2_norm, ["REMAINING"])
        last_sp_idx = _find_col(h2_norm, ["LAST INSPECTION SPAIN"])
        last_eu_idx = _find_col(h2_norm, ["LAST INSPECTION EUROPE"])
        for row in r2:
            code = row[code_idx] if code_idx is not None else None
            if code is None:
                continue
            code = str(code).strip().upper()
            l2_map[code] = {
                "operator": row[op_idx] if op_idx is not None else "",
                "done": row[done_idx] if done_idx is not None else None,
                "objective": row[obj_idx] if obj_idx is not None else None,
                "remaining": row[rem_idx] if rem_idx is not None else None,
                "last": (row[last_sp_idx] if last_sp_idx is not None else None) or (row[last_eu_idx] if last_eu_idx is not None else None),
            }

    sana_map: Dict[str, dict] = {}
    if sana_sheet:
        hs, rs = load_sheet(xlsx_bytes, sana_sheet)
        hs_norm = {_norm_header(h): i for i, h in enumerate(hs) if h}
        code_idx = _find_col(hs_norm, ["OACI", "3LC"])
        op_idx = _find_col(hs_norm, ["OPERADOR", "OPERATOR NAME"])
        obj_idx = _find_col(hs_norm, ["INSPECCIONES OBJETIVO", "OBJECTIVE 2026"])
        done_idx = _find_col(hs_norm, ["INSPECCIONES REALIZ", "INSPECTIONS 2026"])
        rem_idx = _find_col(hs_norm, ["FALTANTES", "REMAINING INSPECTIONS"])
        last_idx = _find_col(hs_norm, ["ULTIMA INSPECCION", "LAST INSPECTION DATE"])
        for row in rs:
            code = row[code_idx] if code_idx is not None else None
            if code is None or str(code).strip() == "":
                continue
            code = str(code).strip().upper()
            sana_map[code] = {
                "operator": row[op_idx] if op_idx is not None else "",
                "done": row[done_idx] if done_idx is not None else None,
                "objective": row[obj_idx] if obj_idx is not None else None,
                "remaining": row[rem_idx] if rem_idx is not None else None,
                "last": row[last_idx] if last_idx is not None else None,
            }

    matriculas_map: Dict[str, str] = {}
    if matriculas_sheet:
        hm, rm = load_sheet(xlsx_bytes, matriculas_sheet, max_col=20)
        hm_norm = {_norm_header(h): i for i, h in enumerate(hm) if h}
        reg_idx = _find_col(hm_norm, ["MATRICULA SIN", "MATRICULA", "REGISTRATION"])
        op_idx = _find_col(hm_norm, ["OPERADOR", "OPERATOR"])
        if reg_idx is not None and op_idx is not None:
            for row in rm:
                reg = row[reg_idx]
                op = row[op_idx]
                if reg and op:
                    reg_key = str(reg).strip().upper().replace(" ", "")
                    matriculas_map[reg_key] = str(op).strip()
                    if "-" in str(reg):
                        matriculas_map[str(reg).strip().upper()] = str(op).strip()

    return icao_map, l1_map, l2_map, sana_map, matriculas_map


def choose_effective_operator(prefix3: str, matricula: str, icao_map: Dict[str, str], matriculas_map: Dict[str, str]) -> Tuple[str, str]:
    matricula_norm = str(matricula).strip().upper() if matricula else ""
    matricula_nohyphen = matricula_norm.replace("-", "")
    if matricula_norm and matricula_norm in matriculas_map:
        return matriculas_map[matricula_norm], "Operador matrícula"
    if matricula_nohyphen and matricula_nohyphen in matriculas_map:
        return matriculas_map[matricula_nohyphen], "Operador matrícula"
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
        texto_busqueda = st.text_input("ARCID (texto libre)", "", key="filtro_arcid")
    with col2:
        ades_disponibles = sorted([a for a in result_df["ADES"].dropna().unique().tolist() if a])
        incluir_vacios = result_df["ADES"].isna().any() or (result_df["ADES"] == "").any()
        opciones_ades = (["(vacío)"] if incluir_vacios else []) + ades_disponibles
        ades_sel = st.multiselect("ADES (destino)", opciones_ades, default=[], key="filtro_ades")
    with col3:
        operadores_disponibles = sorted([o for o in result_df["Operador (maestro)"].dropna().unique().tolist() if o])
        operadores_sel = st.multiselect("Operador", operadores_disponibles, default=[], key="filtro_operador")

    col4, col5, col6 = st.columns(3)
    with col4:
        tipos_disponibles = sorted(result_df["Tipo objetivo"].dropna().unique().tolist())
        tipos_sel = st.multiselect("Tipo objetivo", tipos_disponibles, default=tipos_disponibles, key="filtro_tipo")
    with col5:
        restantes_num = pd.to_numeric(result_df["Restantes"], errors="coerce")
        max_restantes = int(restantes_num.max()) if restantes_num.notna().any() else 0
        restantes_range = st.slider("Restantes (rango)", 0, max(max_restantes, 1), (0, max(max_restantes, 1)), key="filtro_restantes")
    with col6:
        fechas_validas = pd.to_datetime(result_df["Última inspección"], errors="coerce")
        min_fecha = fechas_validas.min().date() if fechas_validas.notna().any() else datetime.now().date()
        max_fecha = fechas_validas.max().date() if fechas_validas.notna().any() else datetime.now().date()
        preset_sel = st.selectbox("Última inspección", ["Todas las fechas", "Última semana", "Último mes", "No en la última semana", "No en el último mes", "Rango personalizado"], index=0, key="filtro_preset")
        fecha_range = st.date_input("Rango personalizado de fechas", value=(min_fecha, max_fecha), key="filtro_fecha_range") if preset_sel == "Rango personalizado" else None

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
        pdf_file = st.file_uploader("1. PDF de tráfico (NOP / ARCID)", type=["pdf"], key="pdf_uploader")
    with col2:
        xlsx_file = st.file_uploader("2. Excel maestro (Objetivos SAFA/SACA/SANA)", type=["xlsx"], key="xlsx_uploader")

    run_clicked = st.button("Generar cruce", type="primary", disabled=not (pdf_file and xlsx_file))

    # Streamlit buttons only return True on the exact rerun triggered by the
    # click; any later interaction (a filter widget, the second button below)
    # would otherwise wipe the whole screen. We persist the processed result
    # in st.session_state so it survives reruns.
    if run_clicked:
        pdf_bytes = pdf_file.read()
        xlsx_bytes = xlsx_file.read()
        flights_df = parse_pdf_flights(pdf_bytes)
        maps = build_master_maps(xlsx_bytes)
        result_df = enrich_flights(flights_df, maps)
        raw_candidates_df = extract_raw_arcid_candidates(pdf_bytes)
        expected_total = extract_expected_total(pdf_bytes)
        st.session_state["result_df"] = result_df
        st.session_state["raw_candidates_df"] = raw_candidates_df
        st.session_state["expected_total"] = expected_total
        st.session_state["fecha_str"] = datetime.now().strftime("%Y%m%d")

    if "result_df" not in st.session_state:
        st.info("Sube ambos archivos y pulsa 'Generar cruce' para empezar.")
        return

    result_df = st.session_state["result_df"]
    raw_candidates_df = st.session_state["raw_candidates_df"]
    expected_total = st.session_state["expected_total"]
    fecha_str = st.session_state["fecha_str"]

    st.success(f"Cruce completado con {len(result_df)} vuelos procesados.")
    if expected_total is not None:
        missing = max(expected_total - len(result_df), 0)
        coverage = (len(result_df) / expected_total * 100) if expected_total else 0
        st.caption(f"Cobertura: {len(result_df)} / {expected_total} vuelos ({coverage:.1f}%). Faltantes estimados: {missing}.")
        if missing > 0 and not raw_candidates_df.empty:
            if st.button("Ver vuelos no detectados"):
                st.session_state["show_missing"] = not st.session_state.get("show_missing", False)
            if st.session_state.get("show_missing"):
                detected_arcids = set(result_df["ARCID"].astype(str).str.upper())
                rc = raw_candidates_df.copy()
                rc["ARCID_norm"] = rc["ARCID_guess"].astype(str).str.upper()
                no_detectados = rc[~rc["ARCID_norm"].isin(detected_arcids)]
                if no_detectados.empty:
                    st.success("No se han encontrado vuelos adicionales sin detectar.")
                else:
                    st.markdown(f"**{len(no_detectados)} vuelo(s) presentes en el texto del PDF pero ausentes en la tabla final:**")
                    st.dataframe(no_detectados[["Hora", "ARCID_guess", "parsed_ok"]].rename(columns={"ARCID_guess": "ARCID (detectado en texto crudo)", "parsed_ok": "Se pudo parsear tipo/aeropuertos"}), use_container_width=True)
    else:
        st.caption(f"No se ha podido leer el total declarado de vuelos en el PDF; se muestran los {len(result_df)} vuelos detectados.")

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
