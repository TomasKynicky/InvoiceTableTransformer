import io
import os
import sys
import tempfile
from pathlib import Path

import streamlit as st
from invoice_processor import process_file, export_to_excel

DEFAULT_API_KEY = os.environ.get("GEMINI_API_KEY", "")

st.set_page_config(page_title="Faktury → Excel", page_icon="📄")

st.title("📄 Faktury a účtenky → Excel")
st.markdown("Nahraj PDF faktury nebo účtenky a stáhni je přehledně v Excelu.")

# --- API klíč ---
api_key = st.text_input(
    "Gemini API klíč",
    value=DEFAULT_API_KEY,
    type="password",
)

# --- Upload ---
uploaded_files = st.file_uploader(
    "Nahraj soubory",
    type=["pdf", "jpg", "jpeg", "png", "bmp", "tiff", "gif"],
    accept_multiple_files=True,
)

# --- Processing ---
if st.button("🚀 Zpracovat", disabled=not (api_key and uploaded_files)):

    progress = st.progress(0, text="Start...")

    all_items = []
    first_meta = None

    for i, uf in enumerate(uploaded_files):
        progress.progress(i / len(uploaded_files), text=f"Zpracovávám {uf.name}...")

        # 🔥 uložit temp file (nutné!)
        suffix = Path(uf.name).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uf.read())
            tmp_path = tmp.name

        try:
            data = process_file(tmp_path, api_key)

            if not first_meta:
                first_meta = data

            all_items.extend(data.get("polozky", []))

        except Exception as e:
            st.warning(f"{uf.name}: {e}")

        finally:
            os.remove(tmp_path)

    progress.empty()

    if not all_items:
        st.error("Nepodařilo se extrahovat žádná data.")
        st.stop()

    merged = {
        "cislo_dokladu": first_meta.get("cislo_dokladu", ""),
        "datum": first_meta.get("datum", ""),
        "dodavatel": first_meta.get("dodavatel", {}),
        "polozky": all_items
    }

    # --- Excel ---
    tmp_path = (
        "/tmp/vystup.xlsx"
        if sys.platform != "win32"
        else os.path.join(os.environ.get("TEMP", "."), "vystup.xlsx")
    )

    export_to_excel(merged, tmp_path)

    with open(tmp_path, "rb") as f:
        output_bytes = f.read()

    st.success("Hotovo!")

    st.download_button(
        label="📥 Stáhnout Excel",
        data=output_bytes,
        file_name="vystup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- Preview ---
    with st.expander("🔍 Náhled"):
        import pandas as pd

        try:
            df = pd.read_excel(io.BytesIO(output_bytes))
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.info(f"Náhled není dostupný: {e}")

st.markdown("---")
st.caption("Hotovo 🚀")