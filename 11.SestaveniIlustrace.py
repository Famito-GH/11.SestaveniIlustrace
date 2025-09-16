import pandas as pd
from pptx import Presentation
import os
import sys
import comtypes.client
import logging
import glob
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

# ---------------- LOGGING ---------------- #
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(BASE_DIR, "log.txt")

logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="w"
)

logging.info("===== Spouštím skript 11.SestaveniIlustrace =====")

# ---------------- SOUBORY ---------------- #
excel_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
if not excel_files:
    logging.error("❌ Nebyl nalezen žádný Excel soubor (.xlsx) ve složce.")
    sys.exit(1)

excel_file = excel_files[0]
logging.info(f"Použitý Excel: {excel_file}")

pptx_file = os.path.join(BASE_DIR, "61_ilustrace.pptx")
if not os.path.exists(pptx_file):
    logging.error(f"❌ PowerPoint soubor {pptx_file} nebyl nalezen.")
    sys.exit(1)

output_folder = os.path.join(BASE_DIR, "exported_slides")
os.makedirs(output_folder, exist_ok=True)

# ---------------- NAČTENÍ EXCELU ---------------- #
try:
    df = pd.read_excel(excel_file, header=0)  # první řádek = názvy sloupců
    logging.info(f"Excel úspěšně načten. Sloupce: {list(df.columns)}")
except Exception as e:
    logging.exception(f"❌ Chyba při načítání Excelu: {e}")
    sys.exit(1)


# Normalizace názvů sloupců (odstraní bílé znaky, sjednotí velikost)
df.columns = df.columns.str.strip()

required_columns = [
    "Číslo modelu",
    "Hmotnost (kg)",
    "ŠÍŘKA",
    "VÝŠKA",
    "HLOUBKA",
    "Šířka popruhu",
    "Maximální délka popruhu",
    "Minimální délka popruhu"
]

missing = [col for col in required_columns if col not in df.columns]
if missing:
    logging.error(f"❌ Excel neobsahuje požadované sloupce: {', '.join(missing)}")
    sys.exit(1)

df = df.set_index("Číslo modelu")

# Normalize index to string and strip whitespace
df.index = df.index.map(lambda x: str(x).strip())

# Mapování názvů shape → názvy sloupců (case-sensitive, zachovat styl)
shape_to_column = {
    "váha": "Hmotnost (kg)",
    "šířka": "ŠÍŘKA",
    "výška": "VÝŠKA",
    "hloubka": "HLOUBKA",
    "šířka popruhu": "Šířka popruhu",
    "max. délka popruhu": "Maximální délka popruhu",
    "min. délka popruhu": "Minimální délka popruhu",
    "CisloModelu": "Číslo modelu",  # Pokud je potřeba
}

# Vyber pouze řádky, kde nejsou žádné NaN v požadovaných sloupcích (bez "Číslo modelu", protože je index)
dropna_columns = [col for col in required_columns if col != "Číslo modelu"] + ["Kód"]
valid_rows = df.dropna(subset=dropna_columns)

slides_processed = 0

for idx, row in valid_rows.iterrows():
    kod = str(row["Kód"]).strip()

    prs = Presentation(pptx_file)
    slide = prs.slides[0]

    for shape in slide.shapes:
        shape_name = shape.name
        if shape_name in shape_to_column:
            excel_col = shape_to_column[shape_name]
            if excel_col == "Číslo modelu":
                value = idx
            else:
                value = row[excel_col]
            # Přidej jednotky a formátování
            if shape_name == "váha":
                value_str = f"{value} kg"
            elif shape_name == "hloubka":
                value_str = f"{int(round(float(value)))} cm"
            elif shape_name in ["šířka", "výška", "šířka popruhu", "max. délka popruhu", "min. délka popruhu"]:
                value_str = f"{value} cm"
            else:
                value_str = str(value)
            shape.text = value_str
            logging.info(f"Kód {kod}: shape '{shape_name}' → {value_str}")

            if hasattr(shape, "text_frame") and shape.text_frame is not None:
                # Zarovnání doprava pro šířku popruhu a hloubku
                if shape_name in ["šířka popruhu", "hloubka"]:
                    for paragraph in shape.text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.RIGHT
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Open Sans"
                        run.font.bold = True
                        # Formátování pro rozměry
                        if shape_name in ["šířka", "výška", "hloubka", "šířka popruhu", "max. délka popruhu", "min. délka popruhu"]:
                            run.font.size = Pt(44)
                        elif shape_name == "váha":
                            run.font.size = Pt(28)

    # Export pouze JPG, bez ukládání PPTX
    try:
        # Ulož dočasný soubor pouze pro export (nutné pro COM API)
        temp_pptx = os.path.join(output_folder, f"__temp_{kod}.pptx")
        prs.save(temp_pptx)

        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(temp_pptx, WithWindow=False)

        export_path = os.path.join(output_folder, f"{kod}_20.jpg")
        presentation.Slides[1].Export(export_path, "JPG")

        presentation.Close()
        powerpoint.Quit()
        logging.info(f"✅ Kód {kod}: exportováno do {export_path}")
        slides_processed += 1

        # Smaž dočasný PPTX soubor
        os.remove(temp_pptx)

    except Exception as e:
        logging.exception(f"❌ Chyba při exportu kódu {kod}: {e}")

logging.info(f"Zpracováno {slides_processed} slidů.")
logging.info("===== Konec skriptu =====")
