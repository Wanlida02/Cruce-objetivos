
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

ATYP_PAT = re.compile(
    r'(DA42|A20N|A21N|A320|A321|A319|AT76|B738|B38M|C680|A332|A350|A330|T380|A380|A300|A306|B772|B763|B788|B789)'
)
TIME_PAT = re.compile(r'^(\d{2}:\d{2})A\s*(.*)$')
FLIGHT_LINE_PAT = re.compile(r'^\d{2}:\d{2}A')


def parse_pdf_flights(pdf_bytes):
    """Extracts (Hora, ARCID, Aeronave, ADEP, ADES) from the NOP-style PDF text.

    Each record follows the pattern: HH:MMA ARCID ATYP+REG(5)+ADEP(4)+ADES(4),
    where spacing between tokens is inconsistent in the raw PDF text extraction
    (the aircraft type can appear glued to the ARCID, and the destination
    airport can appear glued to the registration/origin block). We locate the
    aircraft type code first (known, finite list), split the ARCID off
    everything before it, then use FIXED positional slicing on the remainder
    (stripped of spaces) to recover REG / ADEP / ADES reliably, since the
    registration is always 5 chars and airport codes are always 4 chars.
    """
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        raw_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            raw_lines.extend([ln.strip() for ln in txt.splitlines() if ln.strip()])

    merged = []
    current = None
    for ln in raw_lines:
        if FLIGHT_LINE_PAT.match(ln):
            if current:
                merged.append(current)
            current = ln
        elif current:
            current += " " + ln
    if current:
        merged.append(current)

    rows = []
    for ln in merged:
        m = TIME_PAT.match(ln)
        if not m:
            continue
        hora, rest = m.group(1), m.group(2)

        # Strip leading status/flight-state tokens (e.g. "LF", "LU", "LFU")
        # that sometimes appear glued in front of the real ARCID. These are
        # always pure letters with no digits, while a real ARCID always
        # contains at least one digit.
        tokens = rest.split(" ")
        while tokens and re.fullmatch(r'[A-Z]{1,4}', tokens[0]) and not re.search(r'\d', tokens[0]):
            tokens.pop(0)
        rest = " ".join(tokens)

        am = ATYP_PAT.search(rest)
        if not am:
            first_token = rest.split(" ")[0] if rest else ""
            rows.append({"Hora": hora, "ARCID": first_token, "Aeronave": "",
                         "ADEP": "", "ADES": "",
                         "prefix3": first_token[:3]})
            continue

        arcid = rest[:am.start()].strip()
        atyp = am.group(1)
        remainder = rest[am.end():]

        nospace = remainder.replace(" ", "")
        adep = nospace[5:9]
        ades = nospace[9:13]

        rows.append({
            "Hora": hora,
            "ARCID": arcid,
            "Aeronave": atyp,
            "ADEP": adep,
            "ADES": ades,
            "prefix3": re.match(r'^[A-Z]{3}', arcid).group(0) if re.match(r'^[A-Z]{3}', arcid) else arcid[:3],
        })
    return pd.DataFrame(rows)


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

    ws.merge_cells("B2:L2")
    ws["B2"] = f"GCTS - {fecha_str} - Lista de tráfico enriquecida con Objetivos SAFA/SACA/SANA"
    ws["B2"].font = Font(name="Calibri", size=14, bold=True, color="1F3864")

    ws.merge_cells("B3:L3")
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
    centered = {"Hora", "ARCID", "Aeronave", "ADEP", "ADES", "Tipo objetivo",
                "Inspecciones realizadas", "Objetivo 2026", "Restantes", "Última inspección"}

    for i, row in df.iterrows():
        r = header_row + 1 + i
        for j, h in enumerate(headers, start=2):
            val = row[h]
            if pd.isna(val):
                val = None
            c = ws.cell(row=r, column=j, value=val)
            c.font, c.border = body_font, border
            c.alignment = Alignment(horizontal="center", vertical="center") if h in centered \
                else Alignment(horizontal="left", vertical="center", indent=1)

    last_row = header_row + len(df)
    last_col = 1 + len(headers)
    last_col_letter = get_column_letter(last_col)

    widths = {"Hora": 8, "ARCID": 12, "Aeronave": 11, "ADEP": 8, "ADES": 8,
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

        flights_df = parse_pdf_flights(pdf_bytes)
        if flights_df.empty:
            st.error("No se han podido extraer vuelos del PDF. Revisa que el formato sea el habitual de NOP Eurocontrol.")
            st.stop()

        icao_map, l1_map, l2_map, sana_map = build_master_maps(xlsx_bytes)
        result_df = cross_reference(flights_df, icao_map, l1_map, l2_map, sana_map)

        fecha_str = datetime.now().strftime("%d-%m-%Y")
        excel_buf = build_excel(result_df, fecha_str)
        pdf_buf = build_pdf(result_df, fecha_str)

    st.success(f"Cruce completado: {len(result_df)} vuelos procesados.")

    counts = result_df["Tipo objetivo"].value_counts()
    cols = st.columns(len(counts) if len(counts) > 0 else 1)
    for c, (tipo, n) in zip(cols, counts.items()):
        c.metric(tipo, n)

    st.dataframe(result_df, use_container_width=True, height=500)

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            "Descargar Excel enriquecido",
            data=excel_buf,
            file_name=f"GCTS_{fecha_str}_Enriquecido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dcol2:
        st.download_button(
            "Descargar PDF reconstruido",
            data=pdf_buf,
            file_name=f"GCTS_{fecha_str}_Reconstruido.pdf",
            mime="application/pdf",
        )
else:
    st.info("Sube ambos archivos y pulsa 'Generar cruce' para empezar.")
