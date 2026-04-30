import concurrent.futures
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

# Najdi cestu k poppleru
def get_poppler_path():
    possible_paths = [
        "/opt/homebrew/bin",  # macOS Apple Silicon
        "/usr/local/bin",      # macOS Intel
        "/usr/bin",            # Linux
    ]
    for path in possible_paths:
        if os.path.exists(os.path.join(path, "pdftoppm")):
            return path
    return None

try:
    import google.generativeai as genai
except ImportError:
    genai = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


PROMPT = """Z následujícího dokumentu (faktura nebo účtenka) extrahuj:
- Číslo dokumentu nebo faktury
- Datum vystavení nebo datum prodeje (formát: DD.MM.YYYY nebo DD.MM.YYYY, HH:MM:SS pro účtenky)
- Celková částka k zaplacení s DPH (pouze číslo, bez měny)
- Položky faktury/účtenky – pro každou položku uveď:
  * název
  * množství
  * jednotkovou cenu s DPH
  * jednotkovou cenu bez DPH (pokud je na dokladu explicitně uvedena; jinak nech prázdné)
  * celkovou cenu s DPH
  * celkovou cenu bez DPH (pokud je na dokladu explicitně uvedena; jinak nech prázdné)

Odpověz ve formátu JSON bez markdown značek:
{
    "cislo_dokladu": "",
    "datum": "",
    "celkova_castka_s_dph": "",
    "celkova_castka_bez_dph": "",
    "polozky": [
        {
            "nazev": "",
            "mnozstvi": "",
            "cena_za_jednotku_s_dph": "",
            "cena_za_jednotku_bez_dph": "",
            "celkova_cena_s_dph": "",
            "celkova_cena_bez_dph": ""
        }
    ]
}

DŮLEŽITÉ – částka:
- Pokud je číslo jako "6172518087" bez desetinné části, znamená to 6172518,87 Kč (poslední 2 cifry jsou desetinné)
- Pokud je číslo jako "2975.00" nebo "2975" jde o 2975 Kč
- NIKDY neuváděj jako částku celé číslo bez desetinné části pokud existuje desetinná (např. 21273,00)"""


def convert_pdf_to_images(file_input, dpi=200):
    """
    file_input can be:
      - str or Path -> convert_from_path
      - bytes or file-like object -> convert_from_bytes
    Returns list of PIL Images.
    """
    if pdf2image is None:
        raise RuntimeError("pdf2image není nainstalováno.")
    poppler_path = get_poppler_path()
    kwargs = {"dpi": dpi}
    if poppler_path:
        kwargs["poppler_path"] = poppler_path
    if isinstance(file_input, (str, Path)):
        return pdf2image.convert_from_path(str(file_input), **kwargs)
    if isinstance(file_input, bytes):
        return pdf2image.convert_from_bytes(file_input, **kwargs)
    # file-like
    return pdf2image.convert_from_bytes(file_input.read(), **kwargs)


def normalize_amount(value) -> str:
    if not value:
        return ""
    text = str(value).strip()
    result = ""
    has_dot = '.' in text
    has_comma = ',' in text

    if has_dot and not has_comma:
        parts = text.split('.')
        int_part = parts[0]
        dec_part = parts[1][:2] if len(parts) > 1 else '00'
        if int(dec_part) == 0 and len(parts[1]) <= 2:
            result = int_part
        else:
            result = int_part + ',' + dec_part
    elif has_comma:
        parts = text.split(',')
        result = parts[0] + ',' + parts[1][:2]
    else:
        nums = re.findall(r'\d+', text)
        if nums:
            num_str = ''.join(nums)
            if len(num_str) > 2:
                result = num_str[:-2] + ',' + num_str[-2:]
            else:
                result = '0,' + num_str.zfill(2)

    while result.endswith(',00') or result.endswith(',0'):
        if result.endswith(',00'):
            result = result[:-3]
        elif result.endswith(',0'):
            result = result[:-2]

    if result == '':
        return '0'
    return result


def parse_amount(value: str) -> float:
    if not value:
        return 0.0
    clean = value.replace(' ', '').replace(',', '.')
    try:
        return float(clean)
    except Exception:
        return 0.0


def extract_data(image, api_key: str) -> dict:
    if genai is None:
        raise RuntimeError("google-generativeai není nainstalováno.")
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content([PROMPT, image])
    text = response.text.strip()

    print(f"DEBUG Gemini response: {text[:500]}")

    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()

    try:
        data = json.loads(text)
        polozky_raw = data.get("polozky", [])
        polozky = []
        for p in polozky_raw:
            polozky.append({
                "nazev": p.get("nazev", ""),
                "mnozstvi": p.get("mnozstvi", ""),
                "cena_za_jednotku_s_dph": p.get("cena_za_jednotku_s_dph", p.get("cena_za_jednotku", "")),
                "cena_za_jednotku_bez_dph": p.get("cena_za_jednotku_bez_dph", ""),
                "celkova_cena_s_dph": p.get("celkova_cena_s_dph", p.get("celkova_cena", "")),
                "celkova_cena_bez_dph": p.get("celkova_cena_bez_dph", "")
            })
        return {
            "cislo_dokladu": data.get("cislo_dokladu", ""),
            "datum": data.get("datum", ""),
            "celkova_castka_s_dph": data.get("celkova_castka_s_dph", data.get("celkova_castka", "")),
            "celkova_castka_bez_dph": data.get("celkova_castka_bez_dph", ""),
            "polozky": polozky
        }
    except json.JSONDecodeError as e:
        print(f"JSON parse error: {e}, text: {text[:500]}")
        return {"cislo_dokladu": "", "datum": "", "celkova_castka_s_dph": "", "celkova_castka_bez_dph": "", "polozky": [], "error": f"Nepodařilo se parsovat: {text[:200]}"}


def save_to_excel(data: list, output_path: str):
    if openpyxl is None:
        raise RuntimeError("openpyxl není nainstalováno.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Faktury a účtenky"

    headers = [
        "Číslo dokladu", "Den", "Název položky/Popis",
        "Cena za ks bez DPH", "Cena za ks s DPH",
        "ks", "Celkem bez DPH", "Celkem s DPH"
    ]
    ws.append(headers)

    total_bez = 0.0
    total_s = 0.0
    for item in data:
        cislo = item.get("cislo_dokladu", "")
        datum = item.get("datum", "")
        polozky = item.get("polozky", [])

        if polozky:
            for p in polozky:
                nazev = p.get("nazev", "")
                mnozstvi = str(p.get("mnozstvi", "")).strip()
                cena_bez = normalize_amount(p.get("cena_za_jednotku_bez_dph", ""))
                cena_s = normalize_amount(p.get("cena_za_jednotku_s_dph", ""))
                celkem_bez = normalize_amount(p.get("celkova_cena_bez_dph", ""))
                celkem_s = normalize_amount(p.get("celkova_cena_s_dph", ""))

                row = [
                    cislo,
                    datum,
                    nazev,
                    cena_bez,
                    cena_s,
                    mnozstvi,
                    celkem_bez,
                    celkem_s
                ]
                ws.append(row)
                total_bez += parse_amount(celkem_bez)
                total_s += parse_amount(celkem_s)
        else:
            ws.append([cislo, datum, "", "", "", "", "", ""])

    ws.append(["", "", "", "", "", "", "", ""])
    ws.append([
        "", "CELKEM", "", "", "", "",
        f"{total_bez:.2f}".replace('.', ',') if total_bez else "",
        f"{total_s:.2f}".replace('.', ',') if total_s else ""
    ])

    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

    wb.save(output_path)
    return output_path


def _process_single_file(filename, file_input, api_key):
    image_extensions = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif"}
    ext = Path(filename).suffix.lower()
    try:
        if ext == ".pdf":
            images = convert_pdf_to_images(file_input, dpi=150)
        elif ext in image_extensions:
            if isinstance(file_input, (str, Path)):
                images = [Image.open(file_input)]
            elif isinstance(file_input, bytes):
                images = [Image.open(io.BytesIO(file_input))]
            else:
                images = [Image.open(file_input)]
        else:
            return {"cislo_dokladu": filename, "datum": "", "celkova_castka": "", "polozky": [], "error": "Nepodporovaný formát"}

        all_polozky = []
        first_data = None
        for j, img in enumerate(images):
            data = extract_data(img, api_key)
            if j == 0:
                first_data = data
            if data.get("polozky"):
                all_polozky.extend(data.get("polozky", []))

        if first_data:
            first_data["polozky"] = all_polozky
            return first_data
        else:
            return {"cislo_dokladu": filename, "datum": "", "celkova_castka": "", "polozky": []}
    except Exception as e:
        return {"cislo_dokladu": filename, "datum": "", "celkova_castka": "", "polozky": [], "error": str(e)}


def process_files(files, api_key: str, progress_callback=None, max_workers=4):
    """
    files: list of (filename, file_bytes_or_path)
    progress_callback: optional callable(current, total, message)
    max_workers: kolik souborů se zpracovává paralelně (default 4)
    Returns list of dicts compatible with save_to_excel.
    """
    all_data = [None] * len(files)
    total = len(files)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {
            executor.submit(_process_single_file, filename, file_input, api_key): i
            for i, (filename, file_input) in enumerate(files)
        }

        completed = 0
        for future in concurrent.futures.as_completed(future_to_index):
            i = future_to_index[future]
            try:
                all_data[i] = future.result()
            except Exception as e:
                all_data[i] = {"cislo_dokladu": files[i][0], "datum": "", "celkova_castka": "", "polozky": [], "error": str(e)}

            completed += 1
            if progress_callback:
                progress_callback(completed, total, f"Zpracováno {completed}/{total}")

    if progress_callback:
        progress_callback(total, total, "Hotovo")

    return all_data
