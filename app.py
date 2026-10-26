import copy
import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

BASE = Path(__file__).parent
DATA_FILE = BASE / "Data.xlsx"
TEMPLATE = BASE / "Naveen_Job_Card_Empty.xlsx"

FLOORS = [
    ("First Floor Pricing", "FF"),
    ("First Floor with Basement Pricing", "FF+Bsmt"),
    ("Triplex (1st with basement and stilt room)", "Triplex"),
    ("Second Floor Pricing", "SF"),
    ("Third Floor Pricing", "TF"),
    ("Fourth Floor Pricing", "Top"),
    ("Fourth Floor with Terrace Pricing", "Top (Terrace)"),
    ("Duplex (4th with terrace)", "Duplex"),
]
CARD_COLS = ["Date", "Address", "Location", "Property", "Size", "Floor/s", "Price",
             "Facing", "Road", "Source Name", "Phone No.", "Type", "Tron Remarks",
             "Client Reaction"]
EXCEL_COLS = "ABCDEFGHIJKLMN"
FLAGS = [("Others [Corner]", "Corner"), ("Others [Park Facing]", "Park Facing"),
         ("Others [T-Point]", "T-Point"), ("Others [Resale]", "Resale"),
         ("Others [Gated]", "Gated")]


def parse_price(v):
    """Return price in crores as float, or None if unusable."""
    if v is None or pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if re.search(r"sold|hold|full|all building|gf|^g\b|^g-", s, re.I):
        return None
    m = re.search(r"\d+(\.\d+)?", s.replace(",", "."))
    return float(m.group()) if m else None


def first_number(v):
    m = re.search(r"\d+", str(v))
    return int(m.group()) if m else 0


def to_ts(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(v, unit="D")
    if hasattr(v, "year"):
        return pd.Timestamp(v)
    return pd.to_datetime(str(v), dayfirst=True, errors="coerce")


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE)
    df["_size"] = df["Plot Size (Sq Yard)"].map(first_number)
    df["_ts"] = df["Timestamp"].map(to_ts)
    df["_floors"] = df.apply(
        lambda r: [(lab, parse_price(r[col])) for col, lab in FLOORS
                   if parse_price(r[col]) is not None], axis=1)
    return df


def clean(v):
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else v


def make_row(r, floors, visit_date):
    rem = [str(r["Stage of Property"])]
    c = r["Other Info or Comments"]
    if isinstance(c, str) and c.strip():
        rem.append(c.strip())
    for col, lab in FLAGS:
        if r[col] == "Yes":
            rem.append(lab)
    seen, out = set(), []
    for x in rem:
        if x.lower() not in seen:
            seen.add(x.lower())
            out.append(x)
    out = [x for x in out
           if not (x in dict(FLAGS).values()
                   and any(x.lower() in y.lower() and y != x for y in out))]
    phone = r["Source Contact No."]
    if isinstance(phone, float) and not pd.isna(phone):
        phone = int(phone)
    road = r["Road Size"]
    return {
        "Date": visit_date,
        "Address": r["Plot No."],
        "Location": r["Locality"],
        "Property": str(r["Rooms"]).replace(" ", ""),
        "Size": r["Plot Size (Sq Yard)"],
        "Floor/s": ", ".join(f for f, _ in floors),
        "Price": ", ".join("%g" % p for _, p in floors),
        "Facing": clean(r["Plot Facing"]),
        "Road": int(road) if not pd.isna(road) else "",
        "Source Name": r["Source Name"],
        "Phone No.": clean(phone),
        "Type": r["Source of Building"],
        "Tron Remarks": "; ".join(out),
        "Client Reaction": "",
    }


TEAL = PatternFill("solid", fgColor="00897B")


def build_card(rows, header):
    wb = load_workbook(TEMPLATE)
    ws = wb["Naveen"]
    ws["A4"] = f"Client Name:   {header['name']}"
    ws["E4"] = f"Budget Range: {header['budget']}"
    ws["I4"] = f"Client Type:  {header['ctype']}"
    ws["M4"] = f"RM: {header['rm']}"
    ws["A5"] = f"Contact Number: {header['contact']}"
    ws["E5"] = f"Preferred Micro-Market:  {header['markets']}"

    # Teal fill with white bold text: client info band and column headings
    for r in (4, 5, 7):
        for col in EXCEL_COLS:
            cell = ws[f"{col}{r}"]
            cell.fill = TEAL
            cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")

    # The template has rows 8-44 hidden, so unhide every row first
    for r in range(8, max(ws.max_row, 8 + len(rows)) + 1):
        ws.row_dimensions[r].hidden = False
        ws.row_dimensions[r].height = None

    last_styled = 57
    for i, row in enumerate(rows):
        r = 8 + i
        if r > last_styled:
            for col in EXCEL_COLS:
                src, dst = ws[f"{col}{last_styled}"], ws[f"{col}{r}"]
                dst.font = copy.copy(src.font)
                dst.border = copy.copy(src.border)
                dst.alignment = copy.copy(src.alignment)
                dst.fill = copy.copy(src.fill)
        for col, key in zip(EXCEL_COLS, CARD_COLS):
            v = row.get(key, "")
            ws[f"{col}{r}"] = v if v != "" else None
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- UI
st.set_page_config(page_title="Tron Job Card Builder", layout="wide")
st.title("Tron Job Card Builder")

df = load_data()

with st.sidebar:
    st.header("Client details")
    name = st.text_input("Client name")
    budget = st.text_input("Budget range", placeholder="7.5cr")
    ctype = st.text_input("Client type")
    rm = st.text_input("RM")
    contact = st.text_input("Contact number")
    visit_date = st.text_input("Visit date (optional)", placeholder="27/10/2026")

st.subheader("Search")
c1, c2, c3 = st.columns(3)
with c1:
    localities = st.multiselect("Location / area", sorted(df["Locality"].unique()))
    text = st.text_input("Name / plot no. / source name contains")
with c2:
    rooms = st.multiselect("Property", sorted(df["Rooms"].dropna().unique()),
                           default=["4 BHK"])
    stage = st.multiselect("Stage", ["RTM", "U/C"], default=["RTM", "U/C"])
    sources = st.multiselect("Source type", sorted(df["Source of Building"].unique()))
with c3:
    s1, s2 = st.columns(2)
    min_size = s1.number_input("Min size (sq yd)", min_value=0, value=0, step=10)
    max_size = s2.number_input("Max size (sq yd, 0 = any)", min_value=0, value=0, step=10)
    p1, p2 = st.columns(2)
    min_price = p1.number_input("Min price (Cr)", min_value=0.0, value=0.0, step=0.25)
    max_price = p2.number_input("Max price (Cr, 0 = any)", min_value=0.0, value=0.0, step=0.25)
    latest_only = st.checkbox("Latest entry per plot only", value=True)

d = df.copy()
if latest_only:
    d = d.sort_values("_ts", ascending=False).drop_duplicates(["Plot No.", "Locality"])
if localities:
    d = d[d["Locality"].isin(localities)]
if rooms:
    d = d[d["Rooms"].isin(rooms)]
if stage:
    d = d[d["Stage of Property"].isin(stage)]
if sources:
    d = d[d["Source of Building"].isin(sources)]
if min_size:
    d = d[d["_size"] >= min_size]
if max_size:
    d = d[d["_size"] <= max_size]
if text.strip():
    t = text.strip().lower()
    hay = (d["Plot No."].astype(str) + " " + d["Locality"].astype(str) + " "
           + d["Source Name"].astype(str) + " " + d["Blocks"].astype(str))
    d = d[hay.str.lower().str.contains(re.escape(t))]

rows = []
for _, r in d.iterrows():
    fl = [(lab, p) for lab, p in r["_floors"]
          if p >= min_price and (not max_price or p <= max_price)]
    if not fl:
        continue
    rows.append(make_row(r, fl, visit_date))

rows.sort(key=lambda x: (x["Location"], first_number(x["Size"]), str(x["Address"])))

st.write(f"**{len(rows)} matching properties**")
if rows:
    view = pd.DataFrame(rows)
    view.insert(0, "Select", True)
    edited = st.data_editor(
        view, hide_index=True, width="stretch",
        disabled=[c for c in view.columns if c != "Select"],
        column_config={"Select": st.column_config.CheckboxColumn("Include")},
    )
    chosen = [rows[i] for i in edited.index[edited["Select"]]]
    st.write(f"{len(chosen)} selected for the job card")
    if chosen:
        st.download_button(
            "Download filled job card (Excel)",
            data=build_card(chosen, dict(name=name, budget=budget, ctype=ctype, rm=rm,
                                         contact=contact,
                                         markets=", ".join(localities))),
            file_name=f"Job_Card_{name or 'Client'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("No properties match. Loosen the filters.")
