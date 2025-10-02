import pandas as pd
from pptx import Presentation
import os
import sys
import comtypes.client
import logging
import glob
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import tkinter as tk
from tkinter import messagebox, Listbox, MULTIPLE, END

# ---------------- LOGGING ---------------- #
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

log_file = os.path.join(BASE_DIR, "log.txt")

logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="w"
)

logging.info("===== Spouštím skript 11.SestaveniIlustrace ===== " + BASE_DIR)

# ---------------- SOUBORY ---------------- #
excel_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
pptx_files = glob.glob(os.path.join(BASE_DIR, "*.pptx"))

# Nepoužívej os.path.abspath na soubory, které už jsou v BASE_DIR!
excel_file = excel_files[0] if excel_files else None
pptx_file = pptx_files[0] if pptx_files else None

if not excel_file:
    logging.error("❌ Nebyl nalezen žádný Excel soubor (.xlsx) ve složce: " + BASE_DIR)
if not pptx_file:
    logging.error("❌ Nebyl nalezen žádný PowerPoint soubor (.pptx) ve složce: " + BASE_DIR)

output_folder = os.path.join(BASE_DIR, "exported_slides")
os.makedirs(output_folder, exist_ok=True)

# ---------------- NAČTENÍ EXCELU ---------------- #
df = pd.DataFrame()  # Prázdný DataFrame, pokud není k dispozici Excel soubor
try:
    if excel_file:
        df = pd.read_excel(excel_file, header=0)  # první řádek = názvy sloupců
        logging.info(f"Excel úspěšně načten. Sloupce: {list(df.columns)}")
except Exception as e:
    logging.exception(f"❌ Chyba při načítání Excelu: {e}")

# Normalizace názvů sloupců (odstraní bílé znaky, sjednotí velikost)
df.columns = pd.Index([str(col).strip() for col in df.columns])

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

# Kontrola existence sloupce před nastavením indexu
if "Číslo modelu" in df.columns:
    df = df.set_index("Číslo modelu")
    # Normalize index to string and strip whitespace
    df.index = df.index.map(lambda x: str(x).strip())
else:
    logging.error("❌ Excel neobsahuje sloupec 'Číslo modelu'. Další zpracování bylo přeskočeno.")
    # Pokud sloupec neexistuje, valid_rows bude prázdný DataFrame
    valid_rows = pd.DataFrame()
    # ...a případně můžeš zobrazit chybovou hlášku v GUI:
    # tk.Tk().withdraw()
    # messagebox.showerror("Chyba", "Excel neobsahuje sloupec 'Číslo modelu'.")
    # return

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
missing_dropna = [col for col in dropna_columns if col not in df.columns]
if missing_dropna:
    logging.error(f"❌ Nelze filtrovat platné řádky, chybí sloupce: {', '.join(missing_dropna)}")
    valid_rows = pd.DataFrame()
else:
    valid_rows = df.dropna(subset=dropna_columns)

slides_processed = 0

def export_selected_products(selected_kody=None):
    global slides_processed
    slides_processed = 0

    if getattr(sys, 'frozen', False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    log_file = os.path.join(BASE_DIR, "log.txt")
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filemode="w"
    )
    logging.info("===== Spouštím skript 11.SestaveniIlustrace ===== " + BASE_DIR)

    excel_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    pptx_files = glob.glob(os.path.join(BASE_DIR, "*.pptx"))

    excel_file = excel_files[0] if excel_files else None
    pptx_file = pptx_files[0] if pptx_files else None

    # ...existing code...

    try:
        df = pd.read_excel(excel_file, header=0)
        logging.info(f"Excel úspěšně načten. Sloupce: {list(df.columns)}")
    except Exception as e:
        logging.exception(f"❌ Chyba při načítání Excelu: {e}")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", f"Chyba při načítání Excelu: {e}")
        return

    df.columns = pd.Index([str(col).strip() for col in df.columns])

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
        messagebox.showerror("Chyba", f"Excel neobsahuje požadované sloupce: {', '.join(missing)}")
        return

    # Kontrola existence sloupce před nastavením indexu
    if "Číslo modelu" in df.columns:
        df = df.set_index("Číslo modelu")
        df.index = df.index.map(lambda x: str(x).strip())
    else:
        logging.error("❌ Excel neobsahuje sloupec 'Číslo modelu'. Další zpracování bylo přeskočeno.")
        messagebox.showerror("Chyba", "Excel neobsahuje sloupec 'Číslo modelu'.")
        return

    shape_to_column = {
        "váha": "Hmotnost (kg)",
        "šířka": "ŠÍŘKA",
        "výška": "VÝŠKA",
        "hloubka": "HLOUBKA",
        "šířka popruhu": "Šířka popruhu",
        "max. délka popruhu": "Maximální délka popruhu",
        "min. délka popruhu": "Minimální délka popruhu",
        "CisloModelu": "Číslo modelu",
    }

    dropna_columns = [col for col in required_columns if col != "Číslo modelu"] + ["Kód"]
    missing_dropna = [col for col in dropna_columns if col not in df.columns]
    if missing_dropna:
        logging.error(f"❌ Nelze filtrovat platné řádky, chybí sloupce: {', '.join(missing_dropna)}")
        messagebox.showerror("Chyba", f"Nelze filtrovat platné řádky, chybí sloupce: {', '.join(missing_dropna)}")
        return
    valid_rows = df.dropna(subset=dropna_columns)

    # Filtrování podle výběru uživatele
    if selected_kody is not None:
        if "Kód" not in valid_rows.columns:
            logging.error("❌ Excel neobsahuje sloupec 'Kód'.")
            messagebox.showerror("Chyba", "Excel neobsahuje sloupec 'Kód'.")
            return
        valid_rows = valid_rows[valid_rows["Kód"].astype(str).isin(selected_kody)]

    slides_processed = 0

    for idx, row in valid_rows.iterrows():
        kod = str(row["Kód"]).strip() if "Kód" in row else "NEZNÁMÝ"
        prs = Presentation(pptx_file)
        slide = prs.slides[0]

        for shape in slide.shapes:
            shape_name = shape.name
            if shape_name in shape_to_column and shape_to_column[shape_name] in row:
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
    messagebox.showinfo("Hotovo", f"Zpracováno {slides_processed} slidů.\nVýstup: {output_folder}")

def gui_main():
    if getattr(sys, 'frozen', False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    excel_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    excel_file = excel_files[0] if excel_files else None

    kody = []
    def update_products():
        nonlocal kody
        excel_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
        excel_file = excel_files[0] if excel_files else None
        kody = []
        listbox.delete(0, END)
        if excel_file:
            try:
                df = pd.read_excel(excel_file, header=0)
                df.columns = pd.Index([str(col).strip() for col in df.columns])
                if "Kód" in df.columns:
                    kody = df["Kód"].dropna().astype(str).unique().tolist()
                    for k in kody:
                        listbox.insert(END, k)
            except Exception:
                pass

    if excel_file:
        try:
            df = pd.read_excel(excel_file, header=0)
            df.columns = pd.Index([str(col).strip() for col in df.columns])
            if "Kód" in df.columns:
                kody = df["Kód"].dropna().astype(str).unique().tolist()
        except Exception:
            pass

    root = tk.Tk()
    root.title("Sestavení Ilustrace - Export produktů")
    root.geometry("500x470")

    mode_var = tk.IntVar(value=0)  # 0 = všechny, 1 = konkrétní

    frame_mode = tk.Frame(root)
    frame_mode.pack(anchor="w", padx=10, pady=(10,0))

    tk.Label(frame_mode, text="Vyberte režim exportu:").pack(anchor="w")
    tk.Radiobutton(frame_mode, text="Všechny produkty", variable=mode_var, value=0, command=lambda: toggle_listbox()).pack(anchor="w")
    tk.Radiobutton(frame_mode, text="Konkrétní produkty", variable=mode_var, value=1, command=lambda: toggle_listbox()).pack(anchor="w")

    listbox_label = tk.Label(root, text="Vyberte produkty:")
    listbox = Listbox(root, selectmode=MULTIPLE, height=15)
    for k in kody:
        listbox.insert(END, k)

    def toggle_listbox():
        if mode_var.get() == 0:
            listbox_label.pack_forget()
            listbox.pack_forget()
            btn_update.pack_forget()
        else:
            listbox_label.pack(anchor="w", padx=10, pady=(10,0))
            listbox.pack(fill="both", expand=True, padx=10, pady=(0,10))
            btn_update.pack(fill="x", padx=10, pady=(0,10))

    btn_update = tk.Button(root, text="Aktualizovat produkty", command=update_products)
    toggle_listbox()

    def run_export():
        if mode_var.get() == 0:
            export_selected_products(None)
        else:
            selected = [listbox.get(i) for i in listbox.curselection()]
            if not selected:
                messagebox.showerror("Chyba", "Vyberte alespoň jeden produkt.")
                return
            export_selected_products(selected)

    tk.Button(root, text="Exportovat", command=run_export, bg="#4CAF50", fg="white", height=2).pack(fill="x", padx=10, pady=10)
    root.mainloop()

if __name__ == "__main__":
    gui_main()
    sys.exit()

