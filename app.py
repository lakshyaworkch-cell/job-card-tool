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
    """Read a price cell. Returns (list of prices in Cr, display text) or None.

    Handles: 3.25 Cr, 4.2cr, 3.7, "3,50" (decimal comma), "10. 0 Cr",
    "4.45 Cr With half terrace" (note ignored), ranges like "4.50 Cr-5.0 Cr",
    and lists like "6.5, 6.8". Ignores things like "1/4 basement" and
    "sold", "hold", "GF-", "Full building" entries.
    """
    if v is None or pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return [float(v)], "%g" % v
    s = str(v).strip()
    if re.search(r"sold|hold|full|all building|gf|^g\b|^g-", s, re.I):
        return None
    s = re.sub(r"\d+\s*/\s*\d+", " ", s)             # drop fractions like 1/4
    s = re.sub(r"(\d)\.\s+(\d)", r"\1.\2", s)         # "10. 0" -> "10.0"
    if re.fullmatch(r"\s*\d+,\d{1,2}\s*(cr)?\s*", s, re.I):
        s = s.replace(",", ".")                          # "3,50" -> "3.50"
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", s)]
    nums = [n for n in nums if 0.5 <= n <= 500]
    if not nums:
        return None
    is_range = len(nums) > 1 and re.search(r"\d\s*(cr)?\s*(-|to)\s*\d", s, re.I)
    sep = "-" if is_range else "/"
    return nums, sep.join("%g" % n for n in nums)


def first_number(v):
    m = re.search(r"\d+", str(v))
    return int(m.group()) if m else 0


def to_ts(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(v, unit="D")
    if hasattr(v, "year"):
        return pd.Timestamp(v)
    return pd.to_datetime(str(v), dayfirst=True, errors="coerce")


def is_red(cell):
    """True if the cell has a red fill or red text (our 'sold' marker)."""
    def red_rgb(c):
        try:
            if c is None or c.type != "rgb" or not isinstance(c.rgb, str):
                return False
            h = c.rgb[-6:]
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            return r >= 0xC0 and g <= 0x60 and b <= 0x60
        except Exception:
            return False
    if cell.fill is not None and cell.fill.fill_type == "solid" and red_rgb(cell.fill.fgColor):
        return True
    return cell.font is not None and red_rgb(cell.font.color)


SOLD_ROW_MIN_RED_CELLS = 15   # a row with this many red cells counts as "entire row red"


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE)
    df["_size"] = df["Plot Size (Sq Yard)"].map(first_number)
    df["_ts"] = df["Timestamp"].map(to_ts)

    # Read the red colouring straight from the Excel file (pandas can't see colours)
    ws = load_workbook(DATA_FILE).active
    header = [c.value for c in ws[1]]
    col_idx = {name: i + 1 for i, name in enumerate(header)}
    sold_row, red_floor = {}, {}
    for r in range(2, ws.max_row + 1):
        n_red = sum(1 for c in range(1, len(header) + 1) if is_red(ws.cell(r, c)))
        sold_row[r] = n_red >= SOLD_ROW_MIN_RED_CELLS
        red_floor[r] = {lab for col, lab in FLOORS
                        if col in col_idx and is_red(ws.cell(r, col_idx[col]))}

    df["_sold"] = [sold_row.get(i + 2, False) for i in df.index]

    def floors_of(r):
        out = []
        reds = red_floor.get(r.name + 2, set())
        for col, lab in FLOORS:
            p = parse_price(r[col])
            if p:
                out.append((lab, p[0], p[1], lab in reds))
        return out

    df["_floors"] = df.apply(floors_of, axis=1)
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
        "Floor/s": ", ".join(f for f, _, _ in floors),
        "Price": ", ".join(t for _, _, t in floors),
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
st.markdown("""
<style>
:root { --teal:#00897B; --teal-dark:#00695C; --teal-light:#E0F2F1; }
[data-testid="stHeader"] { background: var(--teal); }
[data-testid="stHeader"] * { color:#fff !important; }
.block-container h1 { background:var(--teal); color:#fff !important; padding:0.6rem 1rem;
                      border-radius:8px; }
.block-container h3 { color:#4DB6AC; border-bottom:2px solid var(--teal); }
[data-testid="stSidebar"] { background: var(--teal); }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] label p,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] span { color:#fff !important; }
[data-testid="stSidebar"] input { color:#1f2d2b !important; background:#fff !important; }
.stButton > button, .stDownloadButton > button {
    background:var(--teal) !important; color:#fff !important; border:0; border-radius:6px; }
.stButton > button p, .stDownloadButton > button p { color:#fff !important; }
.stButton > button:hover, .stDownloadButton > button:hover { background:var(--teal-dark) !important; }
span[data-baseweb="tag"] { background-color:var(--teal) !important; color:#fff !important; }
span[data-baseweb="tag"] span { color:#fff !important; }
</style>
""", unsafe_allow_html=True)
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

INITIAL = {"loc": [], "text": "", "facing": [], "rooms": ["4 BHK"],
           "stage": ["RTM", "U/C"], "sources": [], "min_size": 0, "max_size": 0,
           "min_price": 0.0, "max_price": 0.0, "latest_only": True,
           "hide_rows": True, "hide_floors": True}
CLEARED = {**INITIAL, "rooms": [], "stage": []}
for k, v in INITIAL.items():
    st.session_state.setdefault(k, v)


def clear_filters():
    for k, v in CLEARED.items():
        st.session_state[k] = v


h1, h2 = st.columns([6, 1])
h1.subheader("Search")
h2.button("Clear all", on_click=clear_filters, width="stretch")

c1, c2, c3 = st.columns(3)
with c1:
    localities = st.multiselect("Location / area", sorted(df["Locality"].unique()), key="loc")
    text = st.text_input("Name / plot no. / source name contains", key="text")
    facing = st.multiselect("Facing", sorted(df["Plot Facing"].dropna().unique()), key="facing")
with c2:
    rooms = st.multiselect("Property", sorted(df["Rooms"].dropna().unique()), key="rooms")
    stage = st.multiselect("Stage", ["RTM", "U/C"], key="stage")
    sources = st.multiselect("Source type", sorted(df["Source of Building"].unique()),
                             key="sources")
with c3:
    s1, s2 = st.columns(2)
    min_size = s1.number_input("Min size (sq yd)", min_value=0, step=10, key="min_size")
    max_size = s2.number_input("Max size (sq yd, 0 = any)", min_value=0, step=10,
                               key="max_size")
    p1, p2 = st.columns(2)
    min_price = p1.number_input("Min price (Cr)", min_value=0.0, step=0.25, key="min_price")
    max_price = p2.number_input("Max price (Cr, 0 = any)", min_value=0.0, step=0.25,
                                key="max_price")
    latest_only = st.checkbox("Latest entry per plot only", key="latest_only")
    hide_sold_rows = st.checkbox("Hide sold properties (entire row red)", key="hide_rows")
    hide_sold_floors = st.checkbox("Hide sold floors (red price cell)", key="hide_floors")

d = df.copy()
if latest_only:
    d = d.sort_values("_ts", ascending=False).drop_duplicates(["Plot No.", "Locality"])
n_sold_hidden = 0
if hide_sold_rows:
    n_before = len(d)
    d = d[~d["_sold"]]
    n_sold_hidden = n_before - len(d)
if localities:
    d = d[d["Locality"].isin(localities)]
if facing:
    d = d[d["Plot Facing"].isin(facing)]
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
    fl = [(lab, nums, txt) for lab, nums, txt, red in r["_floors"]
          if not (hide_sold_floors and red and not r["_sold"])
          and any(n >= min_price and (not max_price or n <= max_price) for n in nums)]
    if not fl:
        continue
    rows.append(make_row(r, fl, visit_date))

rows.sort(key=lambda x: (x["Location"], first_number(x["Size"]), str(x["Address"])))

st.write(f"**{len(rows)} matching properties**")
if hide_sold_rows and n_sold_hidden:
    st.caption(f"{n_sold_hidden} sold (red) properties are hidden from the whole list.")
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
