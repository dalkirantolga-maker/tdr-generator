from datetime import datetime, date, time
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

TEMPLATE_FILE = Path("tdr_template.xlsx")

st.set_page_config(
    page_title="TDR Generator",
    page_icon="⚓",
    layout="centered",
)

st.title("⚓ TDR Generator")
st.caption("ALPORT Terminal Departure Report - Excel Generator")


def safe_text(value):
    if value is None:
        return ""
    return str(value).strip()


def safe_filename(value):
    text = safe_text(value).upper().replace(" ", "_").replace("/", "-")
    return text or "UNKNOWN"


def combine_datetime(d: date, t: time):
    if not d or not t:
        return ""
    return datetime.combine(d, t).strftime("%d.%m.%Y %H:%M")


def write_if_exists(ws, cell, value):
    value = safe_text(value)
    if value:
        ws[cell] = value


def create_excel(data):
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError("tdr_template.xlsx bulunamadı. app.py ile aynı yerde olmalı.")

    wb = load_workbook(TEMPLATE_FILE)

    # Ana bilgiler MOVE COUNT sayfasına yazılır.
    ws = wb["MOVE COUNT"] if "MOVE COUNT" in wb.sheetnames else wb.active

    # İlk sürüm hücre eşleşmeleri. Gerekirse şablona göre genişletilir.
    write_if_exists(ws, "E3", data["vessel_name"])
    write_if_exists(ws, "E4", data["voyage_no"])
    write_if_exists(ws, "E5", data["flag"])
    write_if_exists(ws, "E6", data["berth"])
    write_if_exists(ws, "E7", data["agent"])
    write_if_exists(ws, "E8", data["arrival"])
    write_if_exists(ws, "E9", data["berthed"])
    write_if_exists(ws, "E10", data["operation_start"])
    write_if_exists(ws, "E11", data["operation_finish"])
    write_if_exists(ws, "E12", data["departure"])

    # Basit move count alanları. Şablondaki hücreleri netleştirince değiştirilebilir.
    write_if_exists(ws, "B16", data["discharge_20"])
    write_if_exists(ws, "C16", data["discharge_40"])
    write_if_exists(ws, "B17", data["loading_20"])
    write_if_exists(ws, "C17", data["loading_40"])
    write_if_exists(ws, "B18", data["shifting"])
    write_if_exists(ws, "B19", data["reefer"])
    write_if_exists(ws, "B20", data["imdg"])
    write_if_exists(ws, "B21", data["damage"])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


with st.expander("Dosya kontrolü", expanded=False):
    if TEMPLATE_FILE.exists():
        st.success("tdr_template.xlsx bulundu.")
    else:
        st.error("tdr_template.xlsx bulunamadı. Dosyayı app.py ile aynı yere yükle.")

with st.form("tdr_form"):
    st.subheader("1) Vessel Information")
    col1, col2 = st.columns(2)
    with col1:
        vessel_name = st.text_input("Vessel Name", "MARTI SPIRIT")
        voyage_no = st.text_input("Voyage No", "625S")
        flag = st.text_input("Flag", "")
        berth = st.text_input("Berth", "")
    with col2:
        agent = st.text_input("Agent", "")
        arrival_date = st.date_input("Arrival Date", value=date.today())
        arrival_time = st.time_input("Arrival Time", value=time(8, 0))
        berthed_date = st.date_input("Berthed Date", value=date.today())
        berthed_time = st.time_input("Berthed Time", value=time(9, 0))

    st.subheader("2) Operation Times")
    col3, col4 = st.columns(2)
    with col3:
        start_date = st.date_input("Operation Start Date", value=date.today())
        start_time = st.time_input("Operation Start Time", value=time(10, 0))
        finish_date = st.date_input("Operation Finish Date", value=date.today())
        finish_time = st.time_input("Operation Finish Time", value=time(18, 0))
    with col4:
        departure_date = st.date_input("Departure Date", value=date.today())
        departure_time = st.time_input("Departure Time", value=time(20, 0))

    st.subheader("3) Move Count")
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        discharge_20 = st.number_input("Discharge 20'", min_value=0, step=1)
        loading_20 = st.number_input("Loading 20'", min_value=0, step=1)
    with col6:
        discharge_40 = st.number_input("Discharge 40'", min_value=0, step=1)
        loading_40 = st.number_input("Loading 40'", min_value=0, step=1)
    with col7:
        shifting = st.number_input("Shifting", min_value=0, step=1)
        reefer = st.number_input("Reefer", min_value=0, step=1)
    with col8:
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
        "arrival": combine_datetime(arrival_date, arrival_time),
        "berthed": combine_datetime(berthed_date, berthed_time),
        "operation_start": combine_datetime(start_date, start_time),
        "operation_finish": combine_datetime(finish_date, finish_time),
        "departure": combine_datetime(departure_date, departure_time),
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
        excel_file = create_excel(data)
        file_name = f"TDR-{safe_filename(vessel_name)}-{safe_filename(voyage_no)}.xlsx"
        st.success("TDR Excel hazır.")
        st.download_button(
            "Excel Dosyasını İndir",
            data=excel_file,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Hata: {exc}")

st.divider()
st.info("Bu hızlı açılan test sürümüdür. Açılış sorunu çözüldükten sonra liste yükleme ve detay sayfalarını ekleyeceğiz.")
