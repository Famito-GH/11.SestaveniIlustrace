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

    # znovu najdi soubory (pro frozen exe i pro script)
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    excel_files = glob.glob(os.path.join(base_dir, "*.xlsx"))
    pptx_files = glob.glob(os.path.join(base_dir, "*.pptx"))

    excel_file = excel_files[0] if excel_files else None
    pptx_file = pptx_files[0] if pptx_files else None

    if not excel_file or not pptx_file:
        logging.error("❌ Nebyl nalezen Excel nebo PowerPoint soubor.")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", "Nebyl nalezen Excel nebo PowerPoint soubor ve složce.")
        return

    # načti Excel
    try:
        df = pd.read_excel(excel_file, header=0)
        df.columns = pd.Index([str(col).strip() for col in df.columns])
        logging.info(f"Excel úspěšně načten. Sloupce: {list(df.columns)}")
    except Exception as e:
        logging.exception(f"❌ Chyba při načítání Excelu: {e}")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", f"Chyba při načítání Excelu: {e}")
        return

    required_columns = [
        "Číslo modelu",
        "Hmotnost (kg)",
        "ŠÍŘKA",
        "VÝŠKA",
        "HLOUBKA",
        "Šířka popruhu",
        "Maximální délka popruhu",
        "Minimální délka popruhu",
        # "Kód" nemusí být v required_columns protože ho používáme pro filtr/pojmenování
    ]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        logging.error(f"❌ Excel neobsahuje požadované sloupce: {', '.join(missing)}")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", f"Excel neobsahuje požadované sloupce: {', '.join(missing)}")
        return

    # normalizuj číslo modelu jako string sloupec
    df["Číslo modelu"] = df["Číslo modelu"].astype(str).str.strip()

    # zkontroluj, jestli máme Kód když se filtruje podle vybraných kódů
    if selected_kody is not None and "Kód" not in df.columns:
        logging.error("❌ Excel neobsahuje sloupec 'Kód', ale uživatel vybral konkrétní produkty.")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", "Excel neobsahuje sloupec 'Kód'.")
        return

    # dropna pro potřebné sloupce (bez 'Číslo modelu')
    dropna_columns = [col for col in required_columns if col != "Číslo modelu"] + (["Kód"] if "Kód" in df.columns else [])
    missing_dropna = [col for col in dropna_columns if col not in df.columns]
    if missing_dropna:
        logging.error(f"❌ Nelze filtrovat platné řádky, chybí sloupce: {', '.join(missing_dropna)}")
        tk.Tk().withdraw()
        messagebox.showerror("Chyba", f"Nelze filtrovat platné řádky, chybí sloupce: {', '.join(missing_dropna)}")
        return

    valid_rows = df.dropna(subset=dropna_columns)

    # pokud uživatel vybral konkrétní kódy, filtruj podle něj
    if selected_kody is not None:
        valid_rows = valid_rows[valid_rows["Kód"].astype(str).isin(selected_kody)]

    if valid_rows.empty:
        logging.info("ℹ️ Žádné validní řádky k exportu.")
        tk.Tk().withdraw()
        messagebox.showinfo("Hotovo", "Nebyl nalezen žádný produkt k exportu.")
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

    # seskupíme podle čísla modelu
    grouped = valid_rows.groupby("Číslo modelu")

    total_to_process = 0
    for model, group in grouped:
        total_to_process += len(group)

    logging.info(f"Začínám export: {total_to_process} produktů ve {len(grouped)} skupinách podle čísla modelu.")

    # Pro každý model najdi odpovídající slide ve vzorovém PPT a pro každý produkt vygeneruj JPG
    for model, group in grouped:
        model_str = str(model).strip()
        logging.info(f"--- Zpracovávám model: {model_str} (položek: {len(group)}) ---")

        # Pro každý produkt (řádek) otevřeme novou instanci Presentation z originálního PPTX,
        # najdeme slide se shape 'CisloModelu' a porovnáme text.
        for _, row in group.iterrows():
            kod = str(row["Kód"]).strip() if "Kód" in row and not pd.isna(row["Kód"]) else "NEZNAMY"
            try:
                prs = Presentation(pptx_file)

                target_slide_idx = None
                # najdi index slide (0-based) které obsahuje CisloModelu = model_str
                for i, slide in enumerate(prs.slides):
                    found = False
                    for shape in slide.shapes:
                        try:
                            # použij preferenčně shape.text pokud dostupné, jinak text_frame
                            slide_text = ""
                            if hasattr(shape, "text") and shape.text is not None:
                                slide_text = str(shape.text).strip()
                            elif hasattr(shape, "text_frame") and shape.text_frame is not None:
                                slide_text = str(shape.text_frame.text).strip()
                        except Exception:
                            slide_text = ""
                        # porovnej jen pokud má shape jméno CisloModelu
                        if getattr(shape, "name", "") == "CisloModelu" and slide_text == model_str:
                            target_slide_idx = i
                            found = True
                            break
                    if found:
                        break

                if target_slide_idx is None:
                    logging.warning(f"⚠️ Šablonový slide pro model '{model_str}' nebyl nalezen. Kód {kod} přeskočen.")
                    continue

                target_slide = prs.slides[target_slide_idx]

                # doplň hodnoty do shapů na nalezeném slidu
                for shape in target_slide.shapes:
                    shape_name = getattr(shape, "name", "")
                    if shape_name in shape_to_column:
                        excel_col = shape_to_column[shape_name]
                        if excel_col == "Číslo modelu":
                            value = model_str
                        else:
                            # pokud sloupec chybí v řádku, vynech
                            value = row.get(excel_col, "")
                        # formátování
                        try:
                            if shape_name == "váha" and value != "":
                                value_str = f"{value} kg"
                            elif shape_name == "hloubka" and value != "":
                                value_str = f"{int(round(float(value)))} cm"
                            elif shape_name in ["šířka", "výška", "šířka popruhu", "max. délka popruhu", "min. délka popruhu"] and value != "":
                                value_str = f"{value} cm"
                            else:
                                value_str = "" if pd.isna(value) else str(value)
                        except Exception:
                            value_str = str(value)

                        # nastav text (pokud shape není textový, propadne se)
                        try:
                            if hasattr(shape, "text"):
                                shape.text = value_str
                            elif hasattr(shape, "text_frame") and shape.text_frame is not None:
                                # vyčisti a nastav první paragraph
                                shape.text_frame.clear()
                                p = shape.text_frame.paragraphs[0]
                                p.text = value_str
                        except Exception as e:
                            logging.debug(f"Nelze nastavit text pro shape '{shape_name}' na slide {target_slide_idx}: {e}")

                        logging.info(f"Kód {kod}: shape '{shape_name}' → {value_str}")

                        # formátování fontu + zarovnání
                        if hasattr(shape, "text_frame") and shape.text_frame is not None:
                            if shape_name in ["šířka popruhu", "hloubka", "váha", "min. délka popruhu"]:
                                for paragraph in shape.text_frame.paragraphs:
                                    paragraph.alignment = PP_ALIGN.RIGHT
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    try:
                                        run.font.name = "Open Sans"
                                        run.font.bold = True
                                        if shape_name in ["šířka", "výška", "hloubka", "šířka popruhu",
                                                          "max. délka popruhu", "min. délka popruhu"]:
                                            run.font.size = Pt(44)
                                        elif shape_name == "váha":
                                            run.font.size = Pt(28)
                                    except Exception:
                                        pass

                # ulož do temp PPTX a exportuj přes COM
                temp_pptx = os.path.abspath(os.path.join(output_folder, f"__temp_{kod}.pptx"))
                prs.save(temp_pptx)

                powerpoint = None
                presentation = None
                try:
                    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
                    powerpoint.Visible = 1
                    presentation = powerpoint.Presentations.Open(temp_pptx, WithWindow=False)

                    export_path = os.path.abspath(os.path.join(output_folder, f"{kod}_20.jpg"))
                    # COM kolekce Slides je 1-based -> použij Item(index)
                    slide_to_export = presentation.Slides.Item(target_slide_idx + 1)
                    slide_to_export.Export(export_path, "JPG")

                    logging.info(f"✅ Kód {kod}: exportováno do {export_path}")
                    slides_processed += 1

                finally:
                    try:
                        if presentation is not None:
                            presentation.Close()
                    except Exception:
                        pass
                    try:
                        if powerpoint is not None:
                            powerpoint.Quit()
                    except Exception:
                        pass

                # smaž temp soubor
                try:
                    os.remove(temp_pptx)
                except Exception:
                    pass

            except Exception as e:
                logging.exception(f"❌ Chyba při zpracování kódu {kod} (model {model_str}): {e}")

    logging.info(f"Zpracováno {slides_processed} slidů.")
    logging.info("===== Konec skriptu =====")
    tk.Tk().withdraw()
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

