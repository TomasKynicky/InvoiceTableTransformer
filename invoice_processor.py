import io
import json
import os
import re
from pathlib import Path

from PIL import Image

try:
    import pdf2image
except ImportError:
    pdf2image = None

try:
    import google.generativeai as genai
except ImportError:
    genai = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


# =========================
# PROMPT
# =========================

PROMPT = """
Extrahuj fakturu do JSON:

{
  "cislo_dokladu": "",
  "datum": "",
  "dodavatel": {
    "nazev": "",
    "ico": ""
  },
  "celkova_castka_s_dph": "",
  "polozky": [
    {
      "nazev": "",
      "mnozstvi": "",
      "mj": "",
      "cena_za_jednotku": "",
      "cena_celkem": "",
      "cena_celkem_bez_dph": "",
      "cena_celkem_s_dph": "",
      "dph": "",
      "osoba": ""
    }
  ]
}

PRAVIDLA:
- osoba = poslední krátké slovo v řádku (např. Hanzal, DK)
- pokud není → "Nerozpoznané jméno"
- cena_celkem_bez_dph = pokud existuje sloupec bez DPH
- cena_celkem_s_dph = pokud existuje sloupec s DPH
- pokud existuje jen jedna → dej ji do cena_celkem
- zvládni multiline položky
"""


# =========================
# UTILS
# =========================

def get_poppler_path():
    for path in ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin"]:
        if os.path.exists(os.path.join(path, "pdftoppm")):
            return path
    return None


def convert_pdf_to_images(file_input, dpi=200):
    if pdf2image is None:
        raise RuntimeError("pdf2image není nainstalováno")

    kwargs = {"dpi": dpi}
    poppler = get_poppler_path()
    if poppler:
        kwargs["poppler_path"] = poppler

    return pdf2image.convert_from_path(str(file_input), **kwargs)


def parse_price(value: str) -> str:
    if not value:
        return ""

    cleaned = re.sub(r"[^\d.,]", "", value)
    cleaned = cleaned.replace(" ", "").replace(",", ".")

    if re.fullmatch(r"\d+", cleaned):
        if len(cleaned) > 2:
            cleaned = f"{cleaned[:-2]}.{cleaned[-2:]}"

    parts = cleaned.split(".")
    if len(parts) > 2:
        cleaned = f"{''.join(parts[:-1])}.{parts[-1]}"

    return cleaned


def safe_float(val: str) -> float:
    try:
        return float(val.replace(",", "."))
    except:
        return 0.0


def extract_person_name(text: str) -> str:
    words = re.findall(r"\b\w+\b", text)
    if not words:
        return "Nerozpoznané jméno"

    last = words[-1]

    if (
        len(last) <= 10
        and last[0].isupper()
        and not re.search(r"\d", last)
    ):
        return last

    return "Nerozpoznané jméno"


def compute_price_with_dph(item):
    """
    🔥 KLÍČOVÁ LOGIKA DPH
    """

    bez = safe_float(item.get("cena_celkem_bez_dph", ""))
    s = safe_float(item.get("cena_celkem_s_dph", ""))

    if s > 0:
        return s

    if bez > 0:
        # fallback 21 %
        return round(bez * 1.21, 2)

    # fallback na obecné pole
    return safe_float(item.get("cena_celkem", ""))


# =========================
# AI
# =========================

def extract_data(image, api_key: str) -> dict:
    genai.configure(api_key=api_key)

    model = genai.GenerativeModel("gemini-3-flash-preview")
    response = model.generate_content([PROMPT, image])

    text = response.text.strip()
    text = text.replace("```json", "").replace("```", "").strip()

    try:
        data = json.loads(text)
    except:
        return {
            "cislo_dokladu": "",
            "datum": "",
            "dodavatel": {"nazev": "", "ico": ""},
            "celkova_castka_s_dph": "",
            "polozky": []
        }

    supplier = data.get("dodavatel", {"nazev": "", "ico": ""})

    polozky = []

    for p in data.get("polozky", []):
        osoba = p.get("osoba") or extract_person_name(p.get("nazev", ""))

        polozky.append({
            "nazev": p.get("nazev", ""),
            "mnozstvi": p.get("mnozstvi", ""),
            "mj": p.get("mj", ""),

            "cena_za_jednotku": parse_price(p.get("cena_za_jednotku", "")),

            "cena_celkem": parse_price(p.get("cena_celkem", "")),
            "cena_celkem_bez_dph": parse_price(p.get("cena_celkem_bez_dph", "")),
            "cena_celkem_s_dph": parse_price(p.get("cena_celkem_s_dph", "")),

            "dph": parse_price(p.get("dph", "")),
            "osoba": osoba,

            "dodavatel_nazev": supplier.get("nazev", ""),
            "dodavatel_ico": supplier.get("ico", "")
        })

    return {
        "cislo_dokladu": data.get("cislo_dokladu", ""),
        "datum": data.get("datum", ""),
        "dodavatel": supplier,
        "celkova_castka_s_dph": parse_price(data.get("celkova_castka_s_dph", "")),
        "polozky": polozky
    }


# =========================
# GROUPING
# =========================

def group_by_person(items):
    grouped = {}
    for item in items:
        name = item.get("osoba") or "Nerozpoznané jméno"
        grouped.setdefault(name, []).append(item)
    return grouped


# =========================
# EXCEL
# =========================

def export_to_excel(data, path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    grouped = group_by_person(data["polozky"])

    summary = []
    grand_total = 0.0

    for person, items in grouped.items():
        ws = wb.create_sheet(title=person[:31] or "Neznámé")

        ws.append([
            "Doklad", "Datum", "Dodavatel", "IČO",
            "Název", "Množství", "MJ",
            "Cena/ks", "Bez DPH", "S DPH"
        ])

        total = 0.0

        for p in items:
            price_with_dph = compute_price_with_dph(p)
            total += price_with_dph

            ws.append([
                data["cislo_dokladu"],
                data["datum"],
                p["dodavatel_nazev"],
                p["dodavatel_ico"],
                p["nazev"],
                p["mnozstvi"],
                p["mj"],
                p["cena_za_jednotku"],
                p["cena_celkem_bez_dph"],
                f"{price_with_dph:.2f}".replace(".", ",")
            ])

        summary.append((person, total))
        grand_total += total

    ws = wb.create_sheet("Přehled")
    ws.append(["Osoba", "Celkem s DPH"])

    for name, total in summary:
        ws.append([name, f"{total:.2f}".replace(".", ",")])

    ws.append(["CELKEM", f"{grand_total:.2f}".replace(".", ",")])

    wb.save(path)


# =========================
# MAIN
# =========================

def process_file(path, api_key):
    images = convert_pdf_to_images(path)

    all_items = []
    first = None

    for i, img in enumerate(images):
        data = extract_data(img, api_key)

        if i == 0:
            first = data

        all_items.extend(data.get("polozky", []))

    if not first:
        return {}

    first["polozky"] = all_items
    return first


if __name__ == "__main__":
    import sys

    pdf = sys.argv[1]
    api_key = sys.argv[2]

    data = process_file(pdf, api_key)

    with open("output.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    export_to_excel(data, "output.xlsx")

    print("DONE")