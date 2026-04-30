import io
import os
import sys
from pathlib import Path

import streamlit as st
from invoice_processor import process_files, save_to_excel

DEFAULT_API_KEY = os.environ.get("GEMINI_API_KEY", "")

st.set_page_config(page_title="Faktury → Excel", page_icon="📄")

st.title("📄 Faktury a účtenky → Excel")
st.markdown("Nahraj PDF faktury nebo účtenky a stáhni je přehledně v Excelu.")

# --- API klíč ---
api_key = st.text_input(
    "Gemini API klíč",
    value=DEFAULT_API_KEY,
    type="password",
    help="Najdeš na https://aistudio.google.com/app/apikey  (proměnná GEMINI_API_KEY se načte automaticky)"
)

# --- Nahrání souborů ---
uploaded_files = st.file_uploader(
    "Nahraj soubory",
    type=["pdf", "jpg", "jpeg", "png", "bmp", "tiff", "gif"],
    accept_multiple_files=True,
    help="Můžeš nahrát najednou více PDF i obrázků."
)

# --- Zpracování ---
if st.button("🚀 Zpracovat", disabled=not (api_key and uploaded_files)):
    if not api_key:
        st.error("Zadej API klíč.")
        st.stop()
    if not uploaded_files:
        st.error("Nahraj alespoň jeden soubor.")
        st.stop()

    progress_bar = st.progress(0, text="Příprava...")

    def on_progress(current, total, message):
        pct = min(current / total, 1.0) if total else 0.0
        progress_bar.progress(pct, text=message)

    files = []
    for uf in uploaded_files:
        uf.seek(0)
        files.append((uf.name, uf.read()))

    try:
        data = process_files(files, api_key=api_key, progress_callback=on_progress)
    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")
        st.stop()
    finally:
        progress_bar.empty()

    # Uložení do Excelu do paměti
    output_bytes = io.BytesIO()
    tmp_path = "/tmp/vystup_streamlit.xlsx" if sys.platform != "win32" else os.path.join(os.environ.get("TEMP", "."), "vystup_streamlit.xlsx")
    save_to_excel(data, tmp_path)
    with open(tmp_path, "rb") as f:
        output_bytes.write(f.read())
    output_bytes.seek(0)

    st.success("Hotovo! Můžeš si stáhnout výsledný soubor.")
    st.download_button(
        label="📥 Stáhnout vystup.xlsx",
        data=output_bytes,
        file_name="vystup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Zobrazení náhledu
    with st.expander("🔍 Náhled dat"):
        import pandas as pd
        try:
            output_bytes.seek(0)
            df = pd.read_excel(output_bytes)
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.info(f"Náhled není dostupný: {e}")

st.markdown("---")
st.caption("🛠️ V případě problémů zkontroluj API klíč a formát nahraných souborů.")
