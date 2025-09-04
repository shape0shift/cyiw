import tkinter as tk
from tkinter import filedialog, scrolledtext, simpledialog
import matplotlib.pyplot as plt
import numpy as np
import math
import re
import scipy.stats as stats
import pandas as pd  # Für Excel-Export

# ===========================
# Konstanten für Polnisch
K_KONSONANTEN = "bcdfghjklłmnńprsśtwyzźż"
G_KONSONANTEN = "BCDFGHJKLŁMNŃPRSŚTWYZŹŻ"
K_VOKALE = "aąeęiouyó"
G_VOKALE = "AĄEĘIOUYÓ"
SATZENDE = ".!?…"
VERSENDE = "|"
ZAHLEN = "0123456789"
SONSTIGES = "„”«»\"'’‚—–-–*¤/(`),;:_[]{}<>@ \r\n\t"

# Ersetzungen – z.B. gerader Apostroph -> typografischer
ERSATZ_TABELLE = {"'": "’"}

# Digraph-Option: Basis oder erweiterte Digraph-Behandlung
DIGRAPH_BASIS = {
    # leer, Standardbehandlung
}
DIGRAPH_ERWEITERT = {
    "ch": "ç", "Ch": "Ç", "CH": "Ç",
    "cz": "č", "Cz": "Č", "CZ": "Č",
    "dz": "ǳ", "Dz": "ǲ", "DZ": "ǲ",
    "dź": "ď", "Dź": "Ď", "DŹ": "Ď",
    "dż": "ǆ", "Dż": "ǅ", "DŻ": "ǅ",
    "rz": "ž", "Rz": "Ž", "RZ": "Ž",
    "sz": "š", "Sz": "Š", "SZ": "Š"
}

# ===========================
def berechne_statistik(text, digraphs=None):
    # Digraph-Ersetzungen durchführen, falls aktiviert
    if digraphs:
        for alt, neu in digraphs.items():
            text = text.replace(alt, neu)

    # weitere Ersetzungen
    for alt, neu in ERSATZ_TABELLE.items():
        text = text.replace(alt, neu)

    # Satz- und Worttrennung
    satzmuster = r'[.!?…|]+'
    saetze_liste = [s.strip() for s in re.split(satzmuster, text) if s.strip()]
    saetze = len(saetze_liste)

    # einfache Worterkennung (inkl. optionalem Apostroph-Bestandteil)
    wortmuster = r'\b\w+(?:[’]\w+)?\b'
    woerter_liste = re.findall(wortmuster, text, flags=re.UNICODE)
    woerter = len(woerter_liste)

    # Silben zählen: 1 (polnischer) Vokal = 1 Silbe, Ausnahmen für Diphthonge
    vokale = K_VOKALE + G_VOKALE
    diphthonge = ["ia","ią","ie","ię","iu","Ia","Ią","Ie","Ię","Iu"]
    def zaehle_silben(wort):
        wort_tmp = wort
        for diph in diphthonge:
            wort_tmp = wort_tmp.replace(diph, "°")  # Diphthonge durch ein Zeichen ersetzen
        return sum(1 for c in wort_tmp if c in vokale or c == "°")

    silben = sum(zaehle_silben(w) for w in woerter_liste)


    # Grapheme zählen (alle Buchstaben ohne Satzzeichen/Leerzeichen/Ziffern)
    text_ohne_zeichen = re.sub(r'[\s' + re.escape(SATZENDE + SONSTIGES) + ']', '', text)
    grapheme = len(text_ohne_zeichen)

    # Längenmaße
    asl = (woerter / saetze) if saetze else 0.0
    awl = (grapheme / woerter) if woerter else 0.0

    # Flesch
    flesch = 206.835 - 1.015 * asl - 84.6 * (silben / woerter) if woerter and saetze else 0.0

    # Amstad
    amstad = 180 - asl - 58.5 * (silben / woerter) if woerter and saetze else 0.0

    # Tuldava
    i_bar = (silben / woerter) if woerter else 0.0
    j_bar = (woerter / saetze) if saetze else 0.0
    tuldava = i_bar * math.log(j_bar) if j_bar > 0 else 0.0

    # LIX (lange Worte >6 Zeichen)
    lange_worte = sum(1 for w in woerter_liste if len(w) > 6)
    lix = asl + (lange_worte / woerter * 100) if woerter and saetze else 0.0

    # Wiener Sachtextformel (WSTF 1–4)
    ms = (sum(1 for w in woerter_liste if sum(1 for c in w if c in vokale) >= 3) / woerter * 100) if woerter else 0.0
    iw = (lange_worte / woerter * 100) if woerter else 0.0
    es = (sum(1 for w in woerter_liste if sum(1 for c in w if c in vokale) == 1) / woerter * 100) if woerter else 0.0

    wstf1 = 0.1935 * ms + 0.1672 * asl + 0.1297 * iw - 0.0327 * es - 0.875
    wstf2 = 0.2007 * ms + 0.1682 * asl + 0.1373 * iw - 2.779
    wstf3 = 0.2963 * ms + 0.1905 * asl - 1.1144
    wstf4 = 0.2656 * asl + 0.2744 * ms - 1.693

    # New Reading Ease (NRE)
    nre = 1.599 * es - 1.015 * asl - 31.517

    # === Gunning-Fog Index ===
    silben_pro_wort = [sum(1 for c in w if c in vokale) for w in woerter_liste]
    lange_worte_gf = sum(1 for spw in silben_pro_wort if spw >= 3)  # dreisilbige+ Wörter
    gunning_fog = 0.4 * (asl + 100 * (lange_worte_gf / woerter)) if woerter else 0.0

    return {
        "Sätze": saetze,
        "Wörter": woerter,
        "Silben": silben,
        "Grapheme": grapheme,
        "ASL": round(asl, 2),
        "AWL": round(awl, 2),
        "Flesch": round(flesch, 2),
        "Amstad": round(amstad, 2),
        "Tuldava": round(tuldava, 2),
        "Lix": round(lix, 2),
        "WSTF1": round(wstf1, 2),
        "WSTF2": round(wstf2, 2),
        "WSTF3": round(wstf3, 2),
        "WSTF4": round(wstf4, 2),
        "NRE": round(nre, 2),
        "GunningFog": round(gunning_fog, 2)
    }

# ===========================
# Tooltip-Klasse
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        # Mausposition nehmen statt bbox()
        x = event.x_root + 10
        y = event.y_root + 10
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("tahoma", "8", "normal")
        )
        label.pack(ipadx=4, ipady=2)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

# ===========================
class MainGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CYIW ⋅ Calculate Your Index Well ⋅ Polish 1.3")
        self.texts = {}
        self.use_digraphs = tk.BooleanVar(value=False)  # Checkbox-Variable
        self.create_widgets()

    def create_widgets(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=6)

        btn_laden = tk.Button(button_frame, text="📰", command=self.lade_datei, font=("Arial", 18), width=3, height=2)
        btn_laden.pack(side='left', padx=6)
        ToolTip(btn_laden, "Textdatei laden (TXT)")

        btn_liniendiagramm = tk.Button(button_frame, text="📈", command=self.zeige_liniendiagramm, font=("Arial", 18), width=3, height=2)
        btn_liniendiagramm.pack(side='left', padx=6)
        ToolTip(btn_liniendiagramm, "Liniendiagramm")

        btn_streudiagramm = tk.Button(button_frame, text="📊", command=self.zeige_streudiagramm, font=("Arial", 18), width=3, height=2)
        btn_streudiagramm.pack(side='left', padx=6)
        ToolTip(btn_streudiagramm, "Streudiagramm")

        btn_korrelation = tk.Button(button_frame, text="🔢", command=self.zeige_korrelation, font=("Arial", 18), width=3, height=2)
        btn_korrelation.pack(side='left', padx=6)
        ToolTip(btn_korrelation, "Korrelationsmatrix")
        
        btn_excel = tk.Button(button_frame, text="🗒️", command=self.export_excel, font=("Arial", 18), width=3, height=2)
        btn_excel.pack(side='left', padx=6)
        ToolTip(btn_excel, "Als Tabelle speichern")
        
        btn_speichern = tk.Button(button_frame, text="📜", command=self.speichere_ausgabe, font=("Arial", 18), width=3, height=2)
        btn_speichern.pack(side='left', padx=6)
        ToolTip(btn_speichern, "Als TXT speichern")
        
        btn_reset = tk.Button(button_frame, text="♻️", command=self.reset_ausgabe, font=("Arial", 18), width=3, height=2)
        btn_reset.pack(side='left', padx=6)
        ToolTip(btn_reset, "Zurücksetzen")


        # Checkbox für Digraphen-Behandlung
        chk_digraph = tk.Checkbutton(button_frame, text="Digraphs", variable=self.use_digraphs, font=("Arial", 12))
        chk_digraph.pack(side='left', padx=10)
        ToolTip(chk_digraph, "Spezielle Digraph-Ersetzung ein-/ausschalten")

        self.ausgabe_text = scrolledtext.ScrolledText(self.root, width=110, height=30)
        self.ausgabe_text.pack(padx=10, pady=8, fill='both', expand=True)

    # ===========================
    def lade_datei(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text files","*.txt"), ("All files","*.*")])
        if filepath:
            with open(filepath, "r", encoding="utf-8") as f:
                text = f.read()
            kapitel = filepath.split("/")[-1]
            self.texts[kapitel] = text
            self.ausgabe_text.insert(tk.END, f"'{kapitel}' geladen.\n")
            self.analysiere_text(kapitel, text)

    def analysiere_text(self, kapitel, text):
        # Digraph-Ersetzungen durchführen, falls aktiviert
        if self.use_digraphs.get():
            ersetzungen = {
                "ch": "ç", "Ch": "Ç", "CH": "Ç",
                "cz": "č", "Cz": "Č", "CZ": "Č",
                "dz": "ǳ", "Dz": "ǲ", "DZ": "ǲ",
                "dź": "ď", "Dź": "Ď",
                "dż": "ǆ", "Dż": "ǅ",
                "rz": "ž", "Rz": "Ž", "RZ": "Ž",
                "sz": "š", "Sz": "Š", "SZ": "Š"
            }
            for alt, neu in ersetzungen.items():
                text = text.replace(alt, neu)

        ergebnisse = berechne_statistik(text)
        self.ausgabe_text.insert(tk.END, f"\nErgebnisse für {kapitel}:\n")
        for k, v in ergebnisse.items():
            self.ausgabe_text.insert(tk.END, f"{k}: {v}\n")
        self.ausgabe_text.insert(tk.END, "-"*40 + "\n")

    def speichere_ausgabe(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".txt")
        if filepath:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(self.ausgabe_text.get("1.0", tk.END))
            self.ausgabe_text.insert(tk.END, f"\nTXT gespeichert: {filepath}\n")

    def export_excel(self):
        if not self.texts:
            return
        daten = []
        for kapitel, text in self.texts.items():
            stats_dict = berechne_statistik(text)
            stats_dict["Text"] = kapitel
            daten.append(stats_dict)
        df = pd.DataFrame(daten)
        cols = ["Text","Sätze","Wörter","Silben","Grapheme","ASL","AWL",
                "Flesch","Amstad","Tuldava","Lix","WSTF1","WSTF2","WSTF3","WSTF4","NRE","GunningFog"]
        cols = [c for c in cols if c in df.columns]
        df = df[cols]
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if filepath:
            try:
                df.to_excel(filepath, index=False)
                self.ausgabe_text.insert(tk.END, f"\nExcel exportiert: {filepath}\n")
            except Exception as e:
                self.ausgabe_text.insert(tk.END, f"\nFehler beim Excel-Export: {e}\n")

    def zeige_liniendiagramm(self):
        if not self.texts:
            return
        kapitel_namen = []
        indices = {"Flesch": [], "Amstad": [], "Tuldava": [], "Lix": [],
                   "WSTF1": [], "WSTF2": [], "WSTF3": [], "WSTF4": [], "NRE": []}
        for kapitel, text in self.texts.items():
            stats_dict = berechne_statistik(text)
            kapitel_namen.append(kapitel)
            for key in indices:
                indices[key].append(stats_dict.get(key, np.nan))
        plt.figure(figsize=(12, 6))
        for key, values in indices.items():
            plt.plot(kapitel_namen, values, marker="o", label=key)
        plt.xticks(rotation=45, ha="right")
        plt.ylabel("Indexwert")
        plt.title("Textschwierigkeit pro Text")
        plt.legend()
        plt.tight_layout()
        plt.show()

    def zeige_streudiagramm(self):
        if len(self.texts) < 2:
            return
        kapitel = list(self.texts.keys())
        index1 = simpledialog.askstring("Streudiagramm", "Index für X-Achse (z.B. Flesch):")
        index2 = simpledialog.askstring("Streudiagramm", "Index für Y-Achse (z.B. NRE):")
        if not index1 or not index2:
            return
        x_vals, y_vals = [], []
        for k, text in self.texts.items():
            stats = berechne_statistik(text)
            if index1 in stats and index2 in stats:
                x_vals.append(stats[index1])
                y_vals.append(stats[index2])
        if not x_vals:
            return
        plt.figure(figsize=(8, 6))
        plt.scatter(x_vals, y_vals)
        for i, txt in enumerate(kapitel):
            plt.annotate(txt, (x_vals[i], y_vals[i]))
        plt.xlabel(index1)
        plt.ylabel(index2)
        plt.title(f"{index1} vs {index2}")
        plt.tight_layout()
        plt.show()

    def zeige_korrelation(self):
        if not self.texts:
            return
        indices = ["Flesch","Amstad","Tuldava","Lix","WSTF1","WSTF2","WSTF3","WSTF4","NRE"]
        data = {key: [] for key in indices}
        for text in self.texts.values():
            stats_dict = berechne_statistik(text)
            for key in indices:
                data[key].append(stats_dict.get(key, np.nan))
        n_texts = len(self.texts)
        self.ausgabe_text.insert(tk.END, "\nKorrelationsmatrix (Pearson):\n")
        self.ausgabe_text.insert(tk.END, " \t" + "\t".join(indices) + "\n")
        for i, key1 in enumerate(indices):
            line_values, line_sigs, line_ns = [], [], []
            for j, key2 in enumerate(indices):
                if i == j:
                    r, p = 1.000, 0.000
                else:
                    try:
                        r, p = stats.pearsonr(np.array(data[key1]), np.array(data[key2]))
                    except Exception:
                        r, p = float('nan'), float('nan')
                line_values.append(f"{r:.3f}" if not np.isnan(r) else "nan")
                line_sigs.append(f"{p:.3f}" if not np.isnan(p) else "nan")
                line_ns.append(str(n_texts))
            self.ausgabe_text.insert(tk.END, key1 + "\t" + "\t".join(line_values) + "\n")
            self.ausgabe_text.insert(tk.END, "Sig. (2-tailed)\t" + "\t".join(line_sigs) + "\n")
            self.ausgabe_text.insert(tk.END, "N\t" + "\t".join(line_ns) + "\n")

    def reset_ausgabe(self):
        # Canvas/Textfeld leeren
        self.ausgabe_text.delete("1.0", tk.END)
        # Gespeicherte Texte löschen
        self.texts.clear()


# ===========================
if __name__ == "__main__":
    root = tk.Tk()
    app = MainGUI(root)
    root.mainloop()
