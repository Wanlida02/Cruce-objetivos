import re

REG_EXPECTED_LEN = {
    'A7': 5,
    'A6': 5,
    'EC': 5,
}

CURRENT_PDF_ICAOS = set()

def _choose_best_reg(reg_raw):
    if not reg_raw:
        return ''
    candidates = []
    for p in REG_PREFIXES_SORTED:
        if reg_raw.startswith(p) and len(reg_raw) > len(p):
            suffix = reg_raw[len(p):]
            if 2 <= len(suffix) <= 4 and suffix.isalnum():
                reg = f"{p}-{suffix}"
                score = 0
                if len(suffix) > 3:
                    score += 1
                exp = REG_EXPECTED_LEN.get(p)
                if exp and len(suffix) == exp - len(p):
                    score -= 1
                candidates.append((score, reg))
    if candidates:
        candidates.sort(key=lambda x: x[0])
        return candidates[0][1]
    return reg_raw

import re
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A3
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

st.set_page_config(page_title="GCTS - Cruce Objetivos SAFA/SACA/SANA", layout="wide")

st.title("Cruce automático: Lista de tráfico (NOP) + Objetivos SAFA/SACA/SANA")
st.caption(
    "Sube el PDF de la lista de tráfico (NOP Eurocontrol, formato ARCID) y el Excel maestro "
    "de Objetivos SAFA/SACA/SANA/Matrículas. La app cruza cada vuelo por el prefijo ARCID (3 letras) "
    "y genera un Excel y un PDF enriquecidos con Tipo de objetivo, inspecciones realizadas, "
    "objetivo 2026, restantes y última inspección."
)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("1. PDF de tráfico (NOP / ARCID)", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("2. Excel maestro (Objetivos SAFA/SACA/SANA)", type=["xlsx"])

run = st.button("Generar cruce", type="primary", disabled=not (pdf_file and xlsx_file))

# ATYP_PAT is a HINT used to disambiguate ARCID parsing, never a mandatory gate
# that drops a flight when it doesn't match. Real NOP traffic mixes airliners
# with a huge variety of business/regional aircraft (Citation, Falcon, Pilatus,
# Learjet, helicopters, etc.), and a closed whitelist can never keep up -- every
# new NOP export tends to surface a type that's missing, silently losing flights.
# It's expanded generously here to maximise disambiguation accuracy, but the
# generic-shape fallback below (looks_like_type / GENERIC_TYPE_PAT) guarantees
# unrecognised types still get parsed instead of dropped.
ATYP_PAT = re.compile(
    r'(DA42|A20N|A21N|A320|A321|A319|A318|AT76|AT75|AT72|B738|B737|B739|B38M|B39M|B37M|'
    r'C680|C68A|C56X|C525|C550|C25A|C25B|C25C|C25M|C650|C500|C510|C560|'
    r'A332|A333|A342|A343|A345|A346|A350|A359|A35K|A330|A339|T380|A380|A300|A306|'
    r'B772|B773|B77L|B77W|B763|B764|B762|B788|B789|B78X|B748|B744|B742|B734|B735|B736|'
    r'E295|E290|E195|E190|E170|E175|E145|E135|E50P|E545|E550|E55P|'
    r'CRJ9|CRJ7|CRJ2|CRJX|BCS1|BCS3|SB20|SF34|'
    r'F900|F2TH|FA7X|FA8X|F7X|GLF4|GLF5|GLF6|GL5T|GL7T|G280|G150|G200|'
    r'LJ35|LJ40|LJ45|LJ60|LJ70|LJ75|H25B|H25C|CL30|CL35|CL60|CL604|CL605|'
    r'PC12|PC24|TBM7|TBM8|TBM9|BE20|BE9L|BE40|R44|R66|EC20|EC30|EC35|EC45|AS50|AS55|A139|A109|'
    r'GLEX|GLF|DHC6|DHC8|SW4|MU2|PA46|PA28|PA31|SR22|SR20|C172|C182|C206|C210|C340|C414|P28U|P180|P68)'
)

# Generic ICAO aircraft-type SHAPE check (used only as a fallback when ATYP_PAT
# doesn't recognise the token): real designators are 2-4 alphanumeric chars,
# start with a letter, and normally contain at least one digit.
GENERIC_TYPE_PAT = re.compile(r'^[A-Z][A-Z0-9]{2,3}')

def _looks_like_type(tok):
    if not tok:
        return False
    for length in (4, 3):
        cand = tok[:length]
        if len(cand) == length and re.match(r'^[A-Z][A-Z0-9]{2,3}$', cand) and re.search(r'\d', cand):
            return True
    return False

# ARCID (call sign) base shape: 2-4 letter airline/operator code + 1-4 digit
# flight number, with an OPTIONAL trailing single letter (radar-label suffix,
# e.g. the "S" in "GES322S" or the "R" in "NJE918R"). Anchoring on this fixed
# shape is far more reliable than anchoring on aircraft type, because ARCID
# always follows it while aircraft type designators are extremely varied.
ARCID_BASE_PAT = re.compile(r'^([A-Z]{2,4}\d{1,4})([A-Z]{0,2})')

def _resolve_arcid(chunk):
    """Splits the ARCID off the front of `chunk` using two strategies, tried
    in order:

    1. REAL SPACE SEPARATOR (most reliable, ~90% of NOP rows): if the source
       PDF text has an actual space between the ARCID and what follows (e.g.
       'GES322S C56XECMSS ...'), and the token before that space matches the
       ARCID shape, use it directly -- no ambiguity possible.

    2. GLUED FALLBACK (~10% of rows, e.g. 'IBE03CYA333ECLUB...' or
       'EZY43WRA319GEZAJ...'): ARCID's trailing radar-label suffix can be
       ZERO, ONE, or TWO letters, and there is no fixed-width rule to know
       which -- it depends on what immediately follows. We try all three
       suffix lengths (2, 1, 0 letters) and pick the LONGEST one whose
       remainder is a recognised aircraft-type token (checked first against
       ATYP_PAT's known designators, then against the generic ICAO-type
       SHAPE as a fallback for unlisted types). Preferring the longest valid
       split avoids leaving stray suffix letters glued onto the type/
       registration block, which corrupts every downstream field.

    Returns (arcid, remainder_after_arcid) or (None, chunk) if no ARCID-shaped
    token is found at all.
    """
    sp = chunk.find(' ')
    if sp != -1 and 2 <= sp <= 8:
        candidate = chunk[:sp]
        if re.match(r'^[A-Z]{2,4}\d{1,4}[A-Z]{0,2}$', candidate):
            return candidate, chunk[sp + 1:]

    m = ARCID_BASE_PAT.match(chunk)
    if not m:
        return None, chunk
    base = m.group(1)
    full_suffix = m.group(2)
    candidates = [(full_suffix[:L], chunk[len(base) + L:]) for L in range(len(full_suffix), -1, -1)]

    for suf, rest in candidates:
        if ATYP_PAT.match(rest):
            return base + suf, rest
    for suf, rest in candidates:
        if _looks_like_type(rest):
            return base + suf, rest
    return base, chunk[len(base):]

TTV_PAT = re.compile(r'[A-Za-z]\s?\d{3}\s?\d{2}-')

# ICAO aircraft nationality (registration) prefixes -- used to reinsert the
# hyphen in registrations that pdfplumber extracts without one, e.g. 'ECMIF' ->
# 'EC-MIF', 'GEUXE' -> 'G-EUXE'. Sorted longest-first so 2-letter prefixes
# (e.g. 'EC', 'EI', 'OE') are tried before 1-letter ones (e.g. 'G', 'D', 'F')
# to avoid misreading a 2-letter-prefix registration as a 1-letter one.
REG_PREFIXES = [
    # 2-letter / 2-char ICAO nationality prefixes (checked first, longest match wins)
    '9H', '9M', '9V', '9A', '9K', '9G', '4X', '4R', '4L',
    'HB', 'HA', 'HS', 'HL', 'HK', 'HP', 'HZ',
    'LX', 'LY', 'LZ', 'LV', 'LN',
    'EC', 'EI', 'EK', 'EP', 'ER', 'ES', 'ET', 'EW', 'EY',
    'OE', 'OO', 'OY', 'OH', 'OK', 'OM', 'OB',
    'PH', 'PK', 'PP', 'PR', 'PT', 'PZ',
    'SE', 'SP', 'ST', 'SU', 'SX',
    'TC', 'TF', 'TG', 'TJ', 'TN', 'TR', 'TS', 'TU', 'TY', 'TZ',
    'UK', 'UR', 'VH', 'VN', 'VP', 'VQ', 'VT',
    'XA', 'XB', 'XC', 'YI', 'YJ', 'YK', 'YL', 'YR', 'YU', 'YV',
    'ZA', 'ZK', 'ZP', 'ZS', 'CN', 'CS', 'CC', 'CP', 'CU', 'CX',
    'A7', 'A6', 'A9', 'A4', '7T', '4X', 'JY',
    # 1-letter prefixes (checked only if no 2-letter prefix matched)
    'G', 'D', 'F', 'N', 'B', 'I', 'J', 'H', 'P', 'V', 'Z', 'C',
]
REG_PREFIXES_SORTED = sorted(set(REG_PREFIXES), key=len, reverse=True)

# Format 1 (simple traffic list): "HH:MMA ARCID ATYP+REG+ADEP ADES" per physical line.
TIME_PAT_F1 = re.compile(r'^(\d{2}:\d{2})A\s*(.*)$')
FLIGHT_LINE_PAT_F1 = re.compile(r'^\d{2}:\d{2}A')

# Format 2 (NM/CFMU detailed list): "HH:MM[A|E|C][status]ARCID ATYP REG ADEP ADES ..." glued
# together with NO space between the time indicator and what follows. Real NOP exports also use
# 'C' as a status indicator (Cancelled/Changed), not just A/E, and pdfplumber sometimes fuses
# two or three flight records into a single extracted text "line" when they share a row of the
# original PDF table. FLIGHT_START_PAT_F2 is used with finditer (not a line-anchored match) so
# every occurrence of "HH:MM[A|E|C]" inside a fused chunk of text is found and parsed separately,
# instead of only the first one on that physical line.
FLIGHT_START_PAT_F2 = re.compile(r'(\d{2}:\d{2})([AEC])(?=\s?(?:[A-Z]{1,4}\s?)?[A-Z]{2,4}\d)')


def _looks_like_format2(text_sample):
    """Format 2 lines start with HH:MM followed directly by A, E or C (no space),
    while Format 1 lines always have a space after the leading 'A' status char."""
    for ln in text_sample.splitlines():
        ln = ln.strip()
        if re.match(r'^\d{2}:\d{2}[AEC]', ln):
            return True
        if re.match(r'^\d{2}:\d{2}A\s', ln):
            return False
    return False


def _parse_format1(raw_lines):
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

    rows = []
    for ln in merged:
        m = TIME_PAT_F1.match(ln)
        if not m:
            continue
        hora, rest = m.group(1), m.group(2)

        tokens = rest.split(" ")
        while tokens and re.fullmatch(r'[A-Z]{1,4}', tokens[0]) and not re.search(r'\d', tokens[0]):
            tokens.pop(0)
        rest = " ".join(tokens)

        am = ATYP_PAT.search(rest)
        if not am:
            first_token = rest.split(" ")[0] if rest else ""
            rows.append({"Hora": hora, "ARCID": first_token, "Aeronave": "",
                         "ADEP": "", "ADES": "", "prefix3": first_token[:3]})
            continue

        arcid = rest[:am.start()].strip()
        atyp = am.group(1)
        remainder = rest[am.end():]
        nospace = remainder.replace(" ", "")
        adep = nospace[5:9]
        ades = nospace[9:13]

        rows.append({
            "Hora": hora, "ARCID": arcid, "Aeronave": atyp, "ADEP": adep, "ADES": ades,
            "prefix3": re.match(r'^[A-Z]{3}', arcid).group(0) if re.match(r'^[A-Z]{3}', arcid) else arcid[:3],
        })
    return pd.DataFrame(rows)


CURRENT_PDF_ICAOS = set()

def _extract_reg_airports(block):
    """Splits `block` (everything between the aircraft type and the Traffic
    Volume anchor) into (registration, ADEP, ADES).

    STRATEGY: prefer REAL SPACES from the source PDF text as the separator --
    they are the most reliable signal available, and work regardless of which
    registration format is in play (European 'EC-MIF', Gulf 'A7-BHS'/'A6-EEI',
    US N-numbers 'N76064', African '7T-VKH', etc. -- these differ wildly in
    shape, so anchoring on registration shape is fragile, but the space before
    an ICAO airport code is consistently present whenever pdfplumber preserved
    it). We pop the last 1-2 whitespace-separated ICAO-shaped (4 pure letters)
    tokens off the end as ADES/ADEP; if the airport code is glued to the
    registration with no space (still common at the ADEP/registration
    boundary), we peel the trailing 4 letters off that same token instead of
    requiring a space there too.

    Falls back to the old whitelist/last-two-4-letter-blocks/fixed-8-char
    approach only when nothing in the block looks like a spaced-off or
    trailing 4-letter airport code at all (fully glued edge case).
    """
    tokens = block.split()

    def _pop_airport(toks):
        if not toks:
            return None
        t = toks[-1]
        if len(t) == 4 and t.isalpha():
            toks.pop()
            return t
        if len(t) > 4 and t[-4:].isalpha():
            toks[-1] = t[:-4]
            return t[-4:]
        return None

    toks = tokens[:]
    ades = _pop_airport(toks)
    adep = _pop_airport(toks) if ades else None

    if ades and adep:
        reg_raw = "".join(toks)
        return reg_raw, adep, ades

    compact = block.replace(" ", "").replace(">", "")
    if not compact:
        return "", "", ""
    airport_hits = [(m.start(), m.group(0)) for m in re.finditer(r'[A-Z]{4}', compact) if m.group(0) in CURRENT_PDF_ICAOS]
    if len(airport_hits) >= 2:
        adep = airport_hits[-2][1]
        ades = airport_hits[-1][1]
        reg_raw = compact[:airport_hits[-2][0]]
        return reg_raw, adep, ades
    generic_hits = list(re.finditer(r'[A-Z]{4}', compact))
    if len(generic_hits) >= 2:
        adep = generic_hits[-2].group(0)
        ades = generic_hits[-1].group(0)
        reg_raw = compact[:generic_hits[-2].start()]
        return reg_raw, adep, ades
    if len(compact) >= 8:
        return compact[:-8], compact[-8:-4], compact[-4:]
    return "", compact[:4], compact[4:8]

def _parse_one_flight_chunk(hora, rest):
    """Parses a single flight's text chunk (already isolated from any neighbouring
    flights that pdfplumber may have fused onto the same physical line) starting
    right after the HH:MM[A|E|C] indicator, e.g.:
    'LU RSC5SK AT76 EC-MIF GCTS GCLP A090 01-06:45+10:45 06:55E fI 06:44aT  10 N    N'

    STRATEGY (whitelist-free, position/shape based): ARCID always sits at the
    very start of this chunk and always follows the fixed shape
    LETTERS+DIGITS(+optional trailing letter), e.g. 'IBE5105', 'GES322S'.
    We anchor on that shape first (via _resolve_arcid) instead of on aircraft
    type, because ARCID shape is fixed while aircraft type designators are
    extremely varied (airliners, regional types, and a huge range of business
    jets/helicopters) and can never be fully covered by a closed whitelist.
    Once the ARCID is split off, the aircraft type is read from the immediate
    remainder using ATYP_PAT as a hint, falling back to a generic ICAO-type
    SHAPE match (letter + 2-3 alphanumeric chars) so unrecognised types are
    still captured instead of causing the whole flight to be dropped.
    """
    rest = rest.lstrip()

    arcid, after_arcid = _resolve_arcid(rest)
    if arcid is None:
        # Chunk doesn't start with an ARCID-shaped token at all (rare noise
        # from pdfplumber) -- fall back to the old aircraft-type-anchored
        # approach as a last resort.
        am = ATYP_PAT.search(rest)
        if not am:
            return None
        prefix = rest[:am.start()]
        atyp = am.group(1)
        remainder = rest[am.end():]
        dm = re.search(r'\d', prefix)
        alpha_run = prefix[:dm.start()] if dm else prefix
        digit_start = dm.start() if dm else len(prefix)
        airline_code = alpha_run[-3:] if len(alpha_run) >= 3 else alpha_run
        arcid = airline_code + prefix[digit_start:]
    else:
        letters_m = re.match(r'^[A-Z]+', arcid)
        airline_code = letters_m.group(0)[-3:] if letters_m else arcid[:3]

        # Locate the Traffic Volume anchor (e.g. "A390 01-") to isolate the
        # ATYP+REG+ADEP+ADES block from everything that follows (times, flags).
        anchor = TTV_PAT.search(after_arcid)
        block = after_arcid[:anchor.start()] if anchor else after_arcid
        block = block.lstrip()

        # Read the aircraft type off the front of the block: prefer ATYP_PAT
        # (known designator) but fall back to the generic ICAO-type shape
        # (letter + 2-3 alphanumeric chars, e.g. 'A359', 'C56X', 'PC24',
        # 'F2TH') so unrecognised business-jet/regional types are still
        # captured rather than causing the flight to be silently dropped.
        atyp = ""
        am = ATYP_PAT.match(block)
        if am:
            atyp = am.group(1)
            remainder = block[am.end():]
        else:
            gm = GENERIC_TYPE_PAT.match(block)
            if gm:
                atyp = gm.group(0)
                remainder = block[gm.end():]
            else:
                remainder = block

        anchor2 = TTV_PAT.search(remainder)
        reg_airport_block = remainder[:anchor2.start()] if anchor2 else remainder

        reg_raw, adep, ades = _extract_reg_airports(reg_airport_block)
        reg = _choose_best_reg(reg_raw) if reg_raw else ''

        return {
            "Hora": hora, "ARCID": arcid.strip(), "Aeronave": atyp, "Matricula": reg,
            "ADEP": adep, "ADES": ades,
            "prefix3": airline_code.strip(),
        }

    # Legacy fallback path (ARCID_BASE_PAT didn't match at all).
    anchor = TTV_PAT.search(remainder)
    reg_airport_block = remainder[:anchor.start()] if anchor else remainder
    reg_raw, adep, ades = _extract_reg_airports(reg_airport_block)
    reg = _choose_best_reg(reg_raw) if reg_raw else ''

    return {
        "Hora": hora, "ARCID": arcid.strip(), "Aeronave": atyp, "Matricula": reg,
        "ADEP": adep, "ADES": ades,
        "prefix3": airline_code.strip(),
    }


def _parse_format2(raw_lines):
    """Parses NM/CFMU-style detailed traffic lists. Uses finditer over each raw text
    line to find EVERY "HH:MM[A|E|C]" occurrence, because pdfplumber sometimes extracts
    two or three flight rows from the source PDF table as a single fused text line.
    Splitting on every match (instead of anchoring only at the start of the line with
    re.match) ensures no flight is silently dropped when this fusion happens."""
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
            parsed = _parse_one_flight_chunk(hora, chunk)
            if parsed:
                rows.append(parsed)
    return pd.DataFrame(rows)


EXPECTED_TOTAL_PAT = re.compile(r'-\s*(\d+)\s*Flights', re.IGNORECASE)


def extract_expected_total(pdf_bytes):
    """Reads the PDF header text (e.g. '... - 127 Flights') to know how many
    flights the source system declares, so we can show a coverage indicator
    versus how many we actually managed to parse."""
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            m = EXPECTED_TOTAL_PAT.search(txt)
            if m:
                return int(m.group(1))
    return None


def extract_raw_arcid_candidates(pdf_bytes):
    """Scans the raw PDF text for every 'HH:MM[A|E|C]' flight-start marker, returning
    the count found. Used purely as a diagnostic signal (independent of the actual
    parsed DataFrame) to detect flights that were present in the source text but did
    not make it into the final parsed result, e.g. due to an unrecognised aircraft
    type code or a malformed remainder field."""
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        raw_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            raw_lines.extend([ln.strip() for ln in txt.splitlines() if ln.strip()])

    full_text = "\n".join(raw_lines)
    if _looks_like_format2(full_text):
        candidates = []
        for ln in raw_lines:
            matches = list(FLIGHT_START_PAT_F2.finditer(ln))
            for i, m in enumerate(matches):
                hora = m.group(1)
                chunk_start = m.end()
                chunk_end = matches[i + 1].start() if i + 1 < len(matches) else len(ln)
                chunk = ln[chunk_start:chunk_end].lstrip()
                parsed = _parse_one_flight_chunk(hora, chunk)
                arcid_guess = parsed["ARCID"] if parsed else chunk[:12]
                candidates.append({"Hora": hora, "ARCID_guess": arcid_guess.strip(), "parsed_ok": parsed is not None})
        return pd.DataFrame(candidates)
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
        candidates = []
        for ln in merged:
            m = TIME_PAT_F1.match(ln)
            if not m:
                continue
            hora, rest = m.group(1), m.group(2)
            tokens = rest.split(" ")
            while tokens and re.fullmatch(r'[A-Z]{1,4}', tokens[0]) and not re.search(r'\d', tokens[0]):
                tokens.pop(0)
            rest2 = " ".join(tokens)
            am = ATYP_PAT.search(rest2)
            arcid_guess = rest2[:am.start()].strip() if am else (rest2.split(" ")[0] if rest2 else "")
            candidates.append({"Hora": hora, "ARCID_guess": arcid_guess, "parsed_ok": am is not None})
        return pd.DataFrame(candidates)


def parse_pdf_flights(pdf_bytes):
    """Extracts (Hora, ARCID, Aeronave, ADEP, ADES) from either of the two known
    NOP/CFMU PDF traffic-list layouts, auto-detecting which one applies."""
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        raw_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            raw_lines.extend([ln.strip() for ln in txt.splitlines() if ln.strip()])

    full_text = "\n".join(raw_lines)
    if _looks_like_format2(full_text):
        df = _parse_format2(raw_lines)
    else:
        df = _parse_format1(raw_lines)

    return df


def _load_sheet(xlsx_bytes, sheetname, max_col=60):
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


def fmt_date(v):
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v)


def build_master_maps(xlsx_bytes):
    # ICAO CODE -> operator name lookup
    h_icao, r_icao = _load_sheet(xlsx_bytes, "ICAO CODE", max_col=2)
    icao_map = {}
    for row in r_icao:
        name, code = row[0], row[1]
        if code is None or name is None:
            continue
        c = str(code).strip().upper()
        if c:
            icao_map[c] = str(name).strip()

    # Layer 1 Objectives
    h1, r1 = _load_sheet(xlsx_bytes, "Layer 1 Objectives")
    idx1 = {h: i for i, h in enumerate(h1) if h}
    l1_map = {}
    for r in r1:
        code = r[idx1.get("3LC")] if "3LC" in idx1 else None
        if code is None:
            continue
        code = str(code).strip().upper()
        l1_map[code] = {
            "operator": r[idx1.get("Operator Name")] if "Operator Name" in idx1 else "",
            "done": r[idx1.get("Progress")] if "Progress" in idx1 else None,
            "objective": r[idx1.get("Mean Target")] if "Mean Target" in idx1 else None,
            "remaining": r[idx1.get("Remaining")] if "Remaining" in idx1 else None,
            "last": r[idx1.get("Last inspection")] if "Last inspection" in idx1 else None,
        }

    # Layer 2 Objectives
    h2, r2 = _load_sheet(xlsx_bytes, "Layer 2 Objectives")
    idx2 = {h: i for i, h in enumerate(h2) if h}
    l2_map = {}
    for r in r2:
        code = r[idx2.get("3LC")] if "3LC" in idx2 else None
        if code is None:
            continue
        code = str(code).strip().upper()
        obj_key = next((k for k in idx2 if "OBJECTIVE" in str(k)), None)
        last_sp_key = next((k for k in idx2 if "LAST INSPECTION SPAIN" in str(k)), None)
        last_eu_key = next((k for k in idx2 if "LAST INSPECTION EUROPE" in str(k)), None)
        l2_map[code] = {
            "operator": r[idx2.get("OPERATOR L2")] if "OPERATOR L2" in idx2 else "",
            "done": r[idx2.get("DONE")] if "DONE" in idx2 else None,
            "objective": r[idx2.get(obj_key)] if obj_key else None,
            "remaining": r[idx2.get("REMAINING")] if "REMAINING" in idx2 else None,
            "last": (r[idx2.get(last_sp_key)] if last_sp_key else None) or (r[idx2.get(last_eu_key)] if last_eu_key else None),
        }

    # SANA Objectives
    h3, r3 = _load_sheet(xlsx_bytes, "SANA Objectives")
    idx3 = {h: i for i, h in enumerate(h3) if h}
    sana_map = {}
    for r in r3:
        code = r[idx3.get("OACI")] if "OACI" in idx3 else None
        if code is None:
            continue
        code = str(code).strip().upper()
        obj_key = next((k for k in idx3 if "OBJETIVO 26" in str(k)), None)
        done_key = next((k for k in idx3 if "REALIZ. 26" in str(k)), None)
        sana_map[code] = {
            "operator": r[idx3.get("OPERADOR")] if "OPERADOR" in idx3 else "",
            "done": r[idx3.get(done_key)] if done_key else None,
            "objective": r[idx3.get(obj_key)] if obj_key else None,
            "remaining": r[idx3.get("FALTANTES")] if "FALTANTES" in idx3 else None,
            "last": r[idx3.get("ULTIMA INSPECCION")] if "ULTIMA INSPECCION" in idx3 else None,
        }

    return icao_map, l1_map, l2_map, sana_map


def cross_reference(flights_df, icao_map, l1_map, l2_map, sana_map):
    out = []
    for _, r in flights_df.iterrows():
        p = r["prefix3"]
        op = icao_map.get(p, "")
        tipo, done, obj, rem, last, src = "No encontrado", None, None, None, "", ""

        if p in l1_map:
            d = l1_map[p]
            tipo, src = "Layer 1", "Layer 1 Objectives"
            op = op or d["operator"]
            done, obj, rem, last = d["done"], d["objective"], d["remaining"], fmt_date(d["last"])
        elif p in sana_map:
            d = sana_map[p]
            tipo, src = "SANA", "SANA Objectives"
            op = op or d["operator"]
            done, obj, rem, last = d["done"], d["objective"], d["remaining"], fmt_date(d["last"])
        elif p in l2_map:
            d = l2_map[p]
            tipo, src = "Layer 2", "Layer 2 Objectives"
            op = op or d["operator"]
            done, obj, rem, last = d["done"], d["objective"], d["remaining"], fmt_date(d["last"])

        out.append({
            "Hora": r["Hora"], "ARCID": r["ARCID"], "Aeronave": r["Aeronave"],
            "Matricula": r.get("Matricula", ""),
            "ADEP": r["ADEP"], "ADES": r["ADES"], "Operador (maestro)": op,
            "Tipo objetivo": tipo, "Inspecciones realizadas": done, "Objetivo 2026": obj,
            "Restantes": rem, "Última inspección": last, "Fuente cruce": src,
        })
    return pd.DataFrame(out)


def build_excel(df, fecha_str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Enriquecido"
    ws.column_dimensions["A"].width = 3

    ws.merge_cells("B2:M2")
    ws["B2"] = f"GCTS - {fecha_str} - Lista de tráfico enriquecida con Objetivos SAFA/SACA/SANA"
    ws["B2"].font = Font(name="Calibri", size=14, bold=True, color="1F3864")

    ws.merge_cells("B3:M3")
    ws["B3"] = "Cruce por ARCID (prefijo 3LC) contra Excel maestro Objetivos_SAFA_SACA_SANA_Matriculas."
    ws["B3"].font = Font(name="Calibri", size=9, italic=True, color="595959")

    header_row = 5
    headers = list(df.columns)
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    for j, h in enumerate(headers, start=2):
        c = ws.cell(row=header_row, column=j, value=h)
        c.fill, c.font = hdr_fill, hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 28

    body_font = Font(name="Calibri", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    centered = {"Hora", "ARCID", "Aeronave", "Matricula", "ADEP", "ADES", "Tipo objetivo",
                "Inspecciones realizadas", "Objetivo 2026", "Restantes", "Última inspección"}

    for i, row in df.iterrows():
        r = header_row + 1 + i
        for j, h in enumerate(headers, start=2):
            val = row[h]
            if pd.isna(val):
                val = None
            c = ws.cell(row=r, column=j, value=val)
            c.font, c.border = body_font, border
            c.alignment = (Alignment(horizontal="center", vertical="center") if h in centered
                           else Alignment(horizontal="left", vertical="center", indent=1))

    last_row = header_row + len(df)
    last_col = 1 + len(headers)
    last_col_letter = get_column_letter(last_col)

    widths = {"Hora": 8, "ARCID": 12, "Aeronave": 11, "Matricula": 12, "ADEP": 8, "ADES": 8,
              "Operador (maestro)": 42, "Tipo objetivo": 14, "Inspecciones realizadas": 12,
              "Objetivo 2026": 12, "Restantes": 10, "Última inspección": 16, "Fuente cruce": 20}
    for j, h in enumerate(headers, start=2):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(h, 14)

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


def build_pdf(df, fecha_str):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="S", fontName="Helvetica", fontSize=6.2, leading=6.8))
    styles.add(ParagraphStyle(name="T", fontName="Helvetica", fontSize=5.6, leading=6.0))

    headers = ["Hora", "ARCID", "Aeronave", "ADEP", "ADES", "Operador maestro",
               "Tipo objetivo", "Realiz.", "Objetivo", "Rest.", "Última inspección", "Fuente"]

    parts = [
        Paragraph(f"GCTS {fecha_str} - PDF reconstruido con columnas añadidas", styles["Title"]),
        Spacer(1, 4),
        Paragraph("Cruce principal por prefijo ARCID (3 letras) con Excel maestro.", styles["S"]),
        Spacer(1, 6),
    ]

    col_widths = [16*mm, 22*mm, 18*mm, 16*mm, 16*mm, 58*mm, 24*mm, 14*mm, 14*mm, 14*mm, 24*mm, 28*mm]
    page_size = 28

    for start in range(0, len(df), page_size):
        chunk = df.iloc[start:start + page_size]
        data = [headers]
        for _, x in chunk.iterrows():
            data.append([
                x["Hora"], x["ARCID"], x["Aeronave"], x["ADEP"], x["ADES"],
                Paragraph(str(x["Operador (maestro)"])[:60], styles["T"]), x["Tipo objetivo"],
                "" if pd.isna(x["Inspecciones realizadas"]) else str(x["Inspecciones realizadas"]),
                "" if pd.isna(x["Objetivo 2026"]) else str(x["Objetivo 2026"]),
                "" if pd.isna(x["Restantes"]) else str(x["Restantes"]),
                str(x["Última inspección"])[:16], x["Fuente cruce"],
            ])
        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("LEADING", (0, 0), (-1, -1), 6.5),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#edf3f8")]),
            ("ALIGN", (0, 0), (4, -1), "CENTER"),
            ("ALIGN", (6, 0), (9, -1), "CENTER"),
        ]))
        parts.append(t)
        if start + page_size < len(df):
            parts.append(PageBreak())

    buf = BytesIO()
    SimpleDocTemplate(buf, pagesize=landscape(A3), leftMargin=8*mm, rightMargin=8*mm,
                       topMargin=8*mm, bottomMargin=8*mm).build(parts)
    buf.seek(0)
    return buf


if run:
    with st.spinner("Leyendo PDF y Excel, y cruzando datos..."):
        pdf_bytes = pdf_file.read()
        xlsx_bytes = xlsx_file.read()

        expected_total = extract_expected_total(pdf_bytes)

        flights_df = parse_pdf_flights(pdf_bytes)
        if flights_df.empty:
            st.error("No se han podido extraer vuelos del PDF. Revisa que el formato sea el habitual de NOP Eurocontrol.")
            st.stop()

        icao_map, l1_map, l2_map, sana_map = build_master_maps(xlsx_bytes)
        result_df = cross_reference(flights_df, icao_map, l1_map, l2_map, sana_map)
        fecha_str = datetime.now().strftime("%d-%m-%Y")
        excel_buf = build_excel(result_df, fecha_str)

        # Also compute the raw-text diagnostic pass (independent of the final DataFrame)
        # so we can offer a "vuelos no detectados" comparison button afterwards.
        raw_candidates_df = extract_raw_arcid_candidates(pdf_bytes)

        # Persist results in session_state so that interacting with filter
        # widgets, or downloading a file (which also triggers a Streamlit
        # rerun / navigation on some browsers, e.g. inline PDF viewers),
        # does NOT discard the cross-reference results and force the user
        # to click "Generar cruce" again. session_state survives reruns
        # within the same browser session/tab.
        st.session_state["result_df"] = result_df
        st.session_state["excel_buf"] = excel_buf.getvalue()
        st.session_state["fecha_str"] = fecha_str
        st.session_state["expected_total"] = expected_total
        st.session_state["raw_candidates_df"] = raw_candidates_df

if "result_df" in st.session_state:
    result_df = st.session_state["result_df"]
    fecha_str = st.session_state["fecha_str"]
    expected_total = st.session_state.get("expected_total")
    raw_candidates_df = st.session_state.get("raw_candidates_df")
    excel_buf = BytesIO(st.session_state["excel_buf"])
    pdf_buf = build_pdf(result_df, fecha_str)

    # Anchor so that after opening a downloaded/printed PDF in a new tab
    # (which some browsers do for inline PDF viewers), the user can click
    # a link that scrolls straight back to the results table instead of
    # having to regenerate the cross-reference.
    st.markdown('<a name="resultados"></a>', unsafe_allow_html=True)

    st.success(f"Cruce completado: {len(result_df)} vuelos procesados.")

    if expected_total:
        detected = len(result_df)
        pct = detected / expected_total * 100 if expected_total else 0
        missing = max(expected_total - detected, 0)
        if missing == 0:
            st.info(f"✅ Cobertura: {detected} de {expected_total} vuelos detectados ({pct:.0f}%). Coincide con el total declarado en el PDF.")
        else:
            st.warning(
                f"⚠️ Cobertura: {detected} de {expected_total} vuelos detectados ({pct:.0f}%). "
                f"Faltan {missing} vuelo(s) por identificar; revisa manualmente el PDF original para esos casos."
            )

        # Comparison button: show which flights were found in the raw PDF text
        # (via the HH:MM[A|E|C] marker scan) but did not end up in the final
        # parsed/cross-referenced table.
        if missing > 0 and raw_candidates_df is not None and not raw_candidates_df.empty:
            if st.button("Ver vuelos no detectados"):
                detected_arcids = set(result_df["ARCID"].astype(str).str.upper())
                raw_candidates_df["ARCID_norm"] = raw_candidates_df["ARCID_guess"].astype(str).str.upper()
                no_detectados = raw_candidates_df[~raw_candidates_df["ARCID_norm"].isin(detected_arcids)]
                if no_detectados.empty:
                    st.success("No se han encontrado vuelos adicionales sin detectar (la diferencia puede deberse a duplicados o formato de hora).")
                else:
                    st.markdown(f"**{len(no_detectados)} vuelo(s) presentes en el texto del PDF pero ausentes en la tabla final:**")
                    st.dataframe(
                        no_detectados[["Hora", "ARCID_guess", "parsed_ok"]].rename(
                            columns={"ARCID_guess": "ARCID (detectado en texto crudo)", "parsed_ok": "Se pudo parsear tipo/aeropuertos"}
                        ),
                        use_container_width=True,
                    )
                    st.caption(
                        "Estos vuelos aparecen en el texto extraído del PDF (marcador de hora + indicador A/E/C) "
                        "pero no llegaron a la tabla final, normalmente porque el tipo de aeronave no coincide con "
                        "los códigos reconocidos (ATYP_PAT) o porque el bloque de texto quedó incompleto."
                    )
    else:
        st.caption(f"No se ha podido leer el total declarado de vuelos en el PDF; se muestran los {len(result_df)} vuelos detectados.")

    counts = result_df["Tipo objetivo"].value_counts()
    cols = st.columns(len(counts) if len(counts) > 0 else 1)
    for c, (tipo, n) in zip(cols, counts.items()):
        c.metric(tipo, n)

    st.markdown("### Filtros antes de descargar")

    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        texto_busqueda = st.text_input("ARCID (texto libre)", "")
    with fcol2:
        ades_disponibles = sorted(
            [a for a in result_df["ADES"].dropna().unique().tolist() if a]
        )
        incluir_ades_vacios = result_df["ADES"].isna().any() or (result_df["ADES"] == "").any()
        opciones_ades = (["(vacío)"] if incluir_ades_vacios else []) + ades_disponibles
        ades_sel = st.multiselect("ADES (destino)", opciones_ades, default=[])
    with fcol3:
        operadores_disponibles = sorted(
            [o for o in result_df["Operador (maestro)"].dropna().unique().tolist() if o]
        )
        operadores_sel = st.multiselect("Operador", operadores_disponibles, default=[])

    fcol4, fcol5, fcol6 = st.columns(3)
    with fcol4:
        tipos_disponibles = sorted(result_df["Tipo objetivo"].dropna().unique().tolist())
        tipos_sel = st.multiselect(
            "Tipo objetivo", tipos_disponibles, default=tipos_disponibles
        )
    with fcol5:
        restantes_num = pd.to_numeric(result_df["Restantes"], errors="coerce")
        max_restantes = int(restantes_num.max()) if restantes_num.notna().any() else 0
        restantes_range = st.slider(
            "Restantes (rango)", 0, max(max_restantes, 1), (0, max(max_restantes, 1))
        )
    with fcol6:
        fechas_validas = pd.to_datetime(result_df["Última inspección"], errors="coerce")
        if fechas_validas.notna().any():
            min_fecha, max_fecha = fechas_validas.min().date(), fechas_validas.max().date()
        else:
            min_fecha = max_fecha = datetime.now().date()

        preset_opciones = [
            "Todas las fechas",
            "Última semana",
            "Último mes",
            "No en la última semana",
            "No en el último mes",
            "Rango personalizado",
        ]
        preset_sel = st.selectbox("Última inspección", preset_opciones, index=0)

        fecha_range = None
        if preset_sel == "Rango personalizado":
            fecha_range = st.date_input(
                "Rango personalizado de fechas", value=(min_fecha, max_fecha)
            )

    filtered_df = result_df.copy()
    if texto_busqueda.strip():
        filtered_df = filtered_df[
            filtered_df["ARCID"].str.contains(texto_busqueda.strip(), case=False, na=False)
        ]
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
    filtered_df = filtered_df[
        rest_num_full.isna() | rest_num_full.between(restantes_range[0], restantes_range[1])
    ]

    hoy = datetime.now().date()
    fechas_filtro = pd.to_datetime(filtered_df["Última inspección"], errors="coerce")

    if preset_sel == "Última semana":
        limite = hoy - pd.Timedelta(days=7)
        mask_fecha = fechas_filtro.notna() & (fechas_filtro.dt.date >= limite) & (fechas_filtro.dt.date <= hoy)
        filtered_df = filtered_df[mask_fecha]
    elif preset_sel == "Último mes":
        limite = hoy - pd.Timedelta(days=30)
        mask_fecha = fechas_filtro.notna() & (fechas_filtro.dt.date >= limite) & (fechas_filtro.dt.date <= hoy)
        filtered_df = filtered_df[mask_fecha]
    elif preset_sel == "No en la última semana":
        limite = hoy - pd.Timedelta(days=7)
        mask_fecha = fechas_filtro.isna() | (fechas_filtro.dt.date < limite)
        filtered_df = filtered_df[mask_fecha]
    elif preset_sel == "No en el último mes":
        limite = hoy - pd.Timedelta(days=30)
        mask_fecha = fechas_filtro.isna() | (fechas_filtro.dt.date < limite)
        filtered_df = filtered_df[mask_fecha]
    elif preset_sel == "Rango personalizado" and isinstance(fecha_range, tuple) and len(fecha_range) == 2:
        mask_fecha = fechas_filtro.isna() | (
            (fechas_filtro.dt.date >= fecha_range[0]) & (fechas_filtro.dt.date <= fecha_range[1])
        )
        filtered_df = filtered_df[mask_fecha]
    # "Todas las fechas" -> no additional filtering

    st.caption(f"Mostrando {len(filtered_df)} de {len(result_df)} vuelos tras aplicar filtros.")
    st.dataframe(filtered_df, use_container_width=True, height=500)

    excel_buf_filtered = build_excel(filtered_df.reset_index(drop=True), fecha_str)
    pdf_buf_filtered = build_pdf(filtered_df.reset_index(drop=True), fecha_str)

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            "Descargar Excel (con filtros aplicados)",
            data=excel_buf_filtered,
            file_name=f"GCTS_{fecha_str}_Enriquecido_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dcol2:
        st.download_button(
            "Descargar PDF (con filtros aplicados)",
            data=pdf_buf_filtered,
            file_name=f"GCTS_{fecha_str}_Reconstruido_filtrado.pdf",
            mime="application/pdf",
        )

    st.markdown("---")
    st.caption("¿Necesitas todo sin filtrar? Descárgalo aquí:")
    dcol3, dcol4 = st.columns(2)
    with dcol3:
        st.download_button(
            "Descargar Excel completo (sin filtros)",
            data=excel_buf,
            file_name=f"GCTS_{fecha_str}_Enriquecido_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dcol4:
        st.download_button(
            "Descargar PDF completo (sin filtros)",
            data=pdf_buf,
            file_name=f"GCTS_{fecha_str}_Reconstruido_completo.pdf",
            mime="application/pdf",
        )
elif not run:
    st.info("Sube ambos archivos y pulsa 'Generar cruce' para empezar.")
