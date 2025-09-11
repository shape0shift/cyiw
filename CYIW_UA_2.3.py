import tkinter as tk
from tkinter import filedialog, scrolledtext, simpledialog
import matplotlib.pyplot as plt
import numpy as np
import math
import re
import scipy.stats as stats
import pandas as pd   # für Excel-Export

# ===========================
# Konstanten für Ukrainisch
K_KONSONANTEN = "бвгґджзйклмнпрстфхцчшщь"
G_KONSONANTEN = "БВГҐДЖЗЙКЛМНПРСТФХЦЧШЩЬ"
K_VOKALE = "аеєиіїоуюя"
G_VOKALE = "АЕЄИІЇОУЮЯ"
SATZENDE = ".!?…"
VERSENDE = "|"
ZAHLEN = "0123456789"
SONSTIGES = "„*¤/(`),;:-_\"'’“«—»[<>]\r\n\t{}"

ERSATZ_TABELLE = {"'": "’"}

# ===========================
def berechne_statistik(text):
    for alt, neu in ERSATZ_TABELLE.items():
        text = text.replace(alt, neu)

    satzmuster = r'[.!?…|]+'
    sätze = re.split(satzmuster, text)
    sätze = [s.strip() for s in sätze if s.strip()]
    saetze = len(sätze)

    wortmuster = r'\b\w+(?:’\w+)?\b'
    woerter_liste = re.findall(wortmuster, text)
    woerter = len(woerter_liste)

    vokale = K_VOKALE + G_VOKALE
    silben = sum(sum(1 for c in w if c in vokale) for w in woerter_liste)

    # Definiere alle Apostroph-Varianten
    APOSTROPHE = "'’‘‛ʻʼ"

    # Grapheme zählen: Buchstaben + eingebettete Apostrophe
    grapheme = 0
    for w in woerter_liste:
        for i, c in enumerate(w):
            if c in K_VOKALE + G_VOKALE + K_KONSONANTEN + G_KONSONANTEN:
                grapheme += 1
            elif c in APOSTROPHE and 0 < i < len(w)-1:  # Apostroph nur in Wortmitte mitzählen
                grapheme += 1

    asl = saetze and woerter / saetze or 0
    awl = woerter and grapheme / woerter or 0


    flesch = 206.835 - 1.015 * asl - 84.6 * (silben / woerter)
    amstad = 180 - asl - 58.5 * (silben / woerter)

    i_bar = silben / woerter
    j_bar = woerter / saetze
    tuldava = i_bar * math.log(j_bar)

    lange_worte = sum(1 for w in woerter_liste if len(w) > 6)
    lix = asl + (lange_worte / woerter * 100)

    ms = sum(1 for w in woerter_liste if sum(1 for c in w if c in vokale) >= 3) / woerter * 100
    iw = lange_worte / woerter * 100
    es = sum(1 for w in woerter_liste if sum(1 for c in w if c in vokale) == 1) / woerter * 100
    wstf1 = 0.1935 * ms + 0.1672 * asl + 0.1297 * iw - 0.0327 * es - 0.875
    wstf2 = 0.2007 * ms + 0.1682 * asl + 0.1373 * iw - 2.779
    wstf3 = 0.2963 * ms + 0.1905 * asl - 1.1144
    wstf4 = 0.2656 * asl + 0.2744 * ms - 1.693

    nosw = sum(1 for w in woerter_liste if sum(1 for c in w if c in vokale) == 1) / woerter * 100
    nre = 1.599 * nosw - 1.015 * asl - 31.517

 # === Gunning-Fog Index ===
    silben_pro_wort = [sum(1 for c in w if c in vokale) for w in woerter_liste]
    lange_worte_gf = sum(1 for spw in silben_pro_wort if spw >= 3)
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
        self.root.title("CYIW ⋅ Calculate Your Index Well ⋅ Ukrainian 2.3")
        # App-Icon setzen
        try:
            icon = tk.PhotoImage(file="cyiw_ua.png")
            self.root.iconphoto(False, icon)
        except Exception as e:
            print(f"Icon konnte nicht geladen werden: {e}")
        self.texts = {}
        self.create_widgets()

    def create_widgets(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        b1 = tk.Button(button_frame, text="📰", command=self.lade_datei, font=("Arial", 20), width=2, height=1)
        b1.pack(side="left", padx=5)
        ToolTip(b1, "TXT laden")

        b3 = tk.Button(button_frame, text="📈", command=self.zeige_liniendiagramm, font=("Arial", 20), width=2, height=1)
        b3.pack(side="left", padx=5)
        ToolTip(b3, "Liniendiagramm")

        b4 = tk.Button(button_frame, text="📊", command=self.zeige_streudiagramm, font=("Arial", 20), width=2, height=1)
        b4.pack(side="left", padx=5)
        ToolTip(b4, "Streudiagramm")

        b5 = tk.Button(button_frame, text="🔢", command=self.zeige_korrelation, font=("Arial", 20), width=2, height=1)
        b5.pack(side="left", padx=5)
        ToolTip(b5, "Korrelationsmatrix")

        b6 = tk.Button(button_frame, text="🗒️", command=self.export_excel, font=("Arial", 20), width=2, height=1)
        b6.pack(side="left", padx=5)
        ToolTip(b6, "Als Tabelle speichern")
        
        b2 = tk.Button(button_frame, text="📜", command=self.speichere_ausgabe, font=("Arial", 20), width=2, height=1)
        b2.pack(side="left", padx=5)
        ToolTip(b2, "Als Text speichern")
        
        btn_reset = tk.Button(button_frame, text="♻️", command=self.reset_ausgabe, font=("Arial", 20), width=2, height=1)
        btn_reset.pack(side='left', padx=5)
        ToolTip(btn_reset, "Zurücksetzen")

        b_info = tk.Button(button_frame, text="ℹ️", command=self.zeige_info, font=("Arial", 20), width=2, height=1)
        b_info.pack(side="left", padx=5)
        ToolTip(b_info, "Formeln & Legende")

        self.ausgabe_text = scrolledtext.ScrolledText(self.root, width=100, height=30)
        self.ausgabe_text.pack(padx=10, pady=10)

    def lade_datei(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text files","*.txt")])
        if filepath:
            with open(filepath,"r",encoding="utf-8") as f:
                text = f.read()
            kapitel = filepath.split("/")[-1]
            self.texts[kapitel] = text
            self.ausgabe_text.insert(tk.END,f"'{kapitel}' geladen.\n")
            self.analysiere_text(kapitel,text)

    def analysiere_text(self, kapitel, text):
        ergebnisse = berechne_statistik(text)
        self.ausgabe_text.insert(tk.END,f"\nErgebnisse für {kapitel}:\n")
        for k,v in ergebnisse.items():
            self.ausgabe_text.insert(tk.END,f"{k}: {v}\n")
        self.ausgabe_text.insert(tk.END,"-"*30+"\n")

    def speichere_ausgabe(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".txt")
        if filepath:
            with open(filepath,"w",encoding="utf-8") as f:
                f.write(self.ausgabe_text.get("1.0", tk.END))

    def export_excel(self):
        if not self.texts:
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if filepath:
            data = {}
            for kapitel, text in self.texts.items():
                data[kapitel] = berechne_statistik(text)
            df = pd.DataFrame(data).T
            df.to_excel(filepath)

    def zeige_liniendiagramm(self):
        if not self.texts:
            return
        kapitel_namen = []
        indices = {"Flesch":[], "Amstad":[], "Tuldava":[], "Lix":[], 
                   "WSTF1":[], "WSTF2":[], "WSTF3":[], "WSTF4":[], "NRE":[]}
        for kapitel, text in self.texts.items():
            stats = berechne_statistik(text)
            kapitel_namen.append(kapitel)
            for key in indices:
                indices[key].append(stats[key])
        plt.figure(figsize=(12,6))
        for key, values in indices.items():
            plt.plot(kapitel_namen, values, marker="o", label=key)
        plt.xticks(rotation=45, ha="right")
        plt.ylabel("Indexwert")
        plt.title("Textschwierigkeit pro Kapitel")
        plt.legend()
        plt.tight_layout()
        plt.show()

    def zeige_streudiagramm(self):
        if len(self.texts)<2:
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
        plt.figure(figsize=(8,6))
        plt.scatter(x_vals, y_vals)
        for i, txt in enumerate(kapitel):
            plt.annotate(txt,(x_vals[i],y_vals[i]))
        plt.xlabel(index1)
        plt.ylabel(index2)
        plt.title(f"{index1} vs {index2}")
        plt.tight_layout()
        plt.show()

    def zeige_korrelation(self):
        if not self.texts:
            return
        indices = ["Flesch","Amstad","Tuldava","Lix","WSTF1","WSTF2","WSTF3","WSTF4","NRE"]
        data = {key:[] for key in indices}
        for text in self.texts.values():
            stats_dict = berechne_statistik(text)
            for key in indices:
                data[key].append(stats_dict[key])
        n_texts = len(self.texts)
        self.ausgabe_text.insert(tk.END, "\nKorrelationsmatrix (Pearson):\n")
        self.ausgabe_text.insert(tk.END, " \t" + "\t".join(indices) + "\n")
        for i, key1 in enumerate(indices):
            line_values = []
            line_sigs = []
            line_ns = []
            for j, key2 in enumerate(indices):
                if i == j:
                    r = 1.0
                    p = 0.0
                else:
                    r, p = stats.pearsonr(data[key1], data[key2])
                line_values.append(f"{r:.3f}")
                line_sigs.append(f"{p:.3f}")
                line_ns.append(str(n_texts))
            self.ausgabe_text.insert(tk.END, key1 + "\t" + "\t".join(line_values) + "\n")
            self.ausgabe_text.insert(tk.END, "Sig. (2-tailed)\t" + "\t".join(line_sigs) + "\n")
            self.ausgabe_text.insert(tk.END, "N\t" + "\t".join(line_ns) + "\n")

    def reset_ausgabe(self):
        # Canvas/Textfeld leeren
        self.ausgabe_text.delete("1.0", tk.END)
        # Gespeicherte Texte löschen
        self.texts.clear()

    def zeige_info(self):
        info_text = f"""
Flesch Reading Ease:
FRE = 206.835 - (1.015 * ASL) - (84.6 * ASW)

FRE für Deutsch (nach Toni Amstad):
FREger = 180 - ASL - ASW * 58.5

Index der objektiven Textschwierigkeit (nach Tuldava):
R = ī * ln(j̄)

Wiener Sachtextformeln:
WSTF1 = 0.1935 * MS + 0.1672 * SL + 0.1297 * IW - 0.0327 * ES -
0.875
WSTF2 = 0.2007 * MS + 0.1682 * SL + 0.1373 * IW - 2.779
WSTF3 = 0.2963 * MS + 0.1905 * SL - 1.1144
WSTF4 = 0.2656 * SL + 0.2744 * MS - 1.693

Lix-Lesbarkeitsindex (nach Carl-Hugo Björnsson):
Lix = SL + LW/GW * 100

FRE für Russisch (nach Andreas Schiestl):
FRErus = 263.625 - 1.015 * ASL - 84.6 * ASW

Gunning-Fog-Index:
GFI = 0.4 * (ASL + 100 * (PCW/GW)

Pisarek-Formeln:
Pisarek_linear = (1/3 * ASL) * (1/3 * PCW) + 1
Pisarek_nonlinear = 1/2 * ((ASL**2 + PCW**2)**0.5)

Legende:
ASL = durchschnittliche Satzlänge (Wörter/Sätze)
ASW = duchschnittliche Silben pro Wort (Silben/Wörter)
ī = mittlere Wortlänge in Silben
j̄ = mittlere Satzlänge in Wörtern
MS = Wörter mit drei oder mehr Silben
SL = mittlere Satzlänge in der Anzahl der Wörter
IW = Prozentanteil der Wörter mit mehr als sechs Buchstaben
ES = Prozentanteil der einsilbigen Wörter
SL = mittlere Satzlänge
LW = Anzahl der Wörter, die aus mehr als sechs Graphemen bestehen
GW = Gesamtzahl aller Wörter
PCW = Prozentanteil der Wörter länger als drei Grapheme
"""
        info_win = tk.Toplevel(self.root)
        info_win.title("Formeln & Legende")
        tk.Label(info_win, text=info_text, justify="left", font=("Arial", 11)).pack(padx=15, pady=15)

# ===========================
if __name__=="__main__":
    root = tk.Tk()
    app = MainGUI(root)
    root.mainloop()
