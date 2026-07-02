from io import BytesIO
from pathlib import Path
from datetime import date, time, datetime

import streamlit as st
from openpyxl import load_workbook, Workbook

st.set_page_config(page_title="TDR Generator", page_icon="⚓", layout="centered")

TEMPLATE_FILE = Path(__file__).parent / "tdr_template.xlsx"


def clean(v):
    return "" if v is None else str(v).strip()


def safe_filename(v):
    return clean(v).upper().replace(" ", "_").replace("/", "-") or "TDR"


def dt(d, t):
    try:
        return datetime.combine(d, t).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return ""


def set_cell(ws, cell, value):
    value = clean(value)
    if value:
        ws[cell] = value


def build_excel(data):
    # If template exists, use it. If not, create a simple workbook so app never crashes.
    if TEMPLATE_FILE.exists():
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb["MOVE COUNT"] if "MOVE COUNT" in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "MOVE COUNT"
        ws["A1"] = "TDR GENERATOR"

    # Safe first-version cell mapping
    cells = {
        "E3": "vessel_name",
        "E4": "voyage_no",
        "E5": "flag",
        "E6": "berth",
        "E7": "agent",
        "E8": "arrival",
        "E9": "berthed",
        "E10": "operation_start",
        "E11": "operation_finish",
        "E12": "departure",
        "B16": "discharge_20",
        "C16": "discharge_40",
        "B17": "loading_20",
        "C17": "loading_40",
        "B18": "shifting",
        "B19": "reefer",
        "B20": "imdg",
        "B21": "damage",
    }
    for cell, key in cells.items():
        try:
            set_cell(ws, cell, data.get(key, ""))
        except Exception:
            pass

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.title("⚓ TDR Generator")
st.caption("Terminal Departure Report Excel Generator")

if TEMPLATE_FILE.exists():
    st.success("tdr_template.xlsx bulundu. Sistem hazır.")
else:
    st.warning("tdr_template.xlsx bulunamadı. Yine de basit Excel oluşturulabilir.")

with st.form("tdr_form"):
    st.subheader("Vessel Information")
    c1, c2 = st.columns(2)
    with c1:
        vessel_name = st.text_input("Vessel Name", "MARTI SPIRIT")
        voyage_no = st.text_input("Voyage No", "625S")
        flag = st.text_input("Flag")
        berth = st.text_input("Berth")
    with c2:
        agent = st.text_input("Agent")
        arrival_date = st.date_input("Arrival Date", value=date.today())
        arrival_time = st.time_input("Arrival Time", value=time(8, 0))
        berthed_date = st.date_input("Berthed Date", value=date.today())
        berthed_time = st.time_input("Berthed Time", value=time(9, 0))

    st.subheader("Operation Times")
    c3, c4 = st.columns(2)
    with c3:
        start_date = st.date_input("Operation Start Date", value=date.today())
        start_time = st.time_input("Operation Start Time", value=time(10, 0))
        finish_date = st.date_input("Operation Finish Date", value=date.today())
        finish_time = st.time_input("Operation Finish Time", value=time(18, 0))
    with c4:
        departure_date = st.date_input("Departure Date", value=date.today())
        departure_time = st.time_input("Departure Time", value=time(20, 0))

    st.subheader("Move Count")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        discharge_20 = st.number_input("Discharge 20'", min_value=0, step=1)
        loading_20 = st.number_input("Loading 20'", min_value=0, step=1)
    with c6:
        discharge_40 = st.number_input("Discharge 40'", min_value=0, step=1)
        loading_40 = st.number_input("Loading 40'", min_value=0, step=1)
    with c7:
        shifting = st.number_input("Shifting", min_value=0, step=1)
        reefer = st.number_input("Reefer", min_value=0, step=1)
    with c8:
        imdg = st.number_input("IMDG", min_value=0, step=1)
        damage = st.number_input("Damage", min_value=0, step=1)

    submitted = st.form_submit_button("TDR Excel Oluştur", type="primary")

if submitted:
    data = {
        "vessel_name": vessel_name,
        "voyage_no": voyage_no,
        "flag": flag,
        "berth": berth,
        "agent": agent,
        "arrival": dt(arrival_date, arrival_time),
        "berthed": dt(berthed_date, berthed_time),
        "operation_start": dt(start_date, start_time),
        "operation_finish": dt(finish_date, finish_time),
        "departure": dt(departure_date, departure_time),
        "discharge_20": discharge_20,
        "discharge_40": discharge_40,
        "loading_20": loading_20,
        "loading_40": loading_40,
        "shifting": shifting,
        "reefer": reefer,
        "imdg": imdg,
        "damage": damage,
    }
    try:
        excel = build_excel(data)
        st.success("TDR Excel hazır.")
        st.download_button(
            "Excel Dosyasını İndir",
            data=excel,
            file_name=f"TDR-{safe_filename(vessel_name)}-{safe_filename(voyage_no)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error("Excel oluşturulurken hata oluştu.")
        st.exception(e)
