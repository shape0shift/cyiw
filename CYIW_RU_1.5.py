import tkinter as tk
from tkinter import filedialog, scrolledtext, simpledialog
import math
import re
import matplotlib.pyplot as plt
import numpy as np
import scipy.stats as stats
import pandas as pd   # für Excel-Export

# ===============================
vokale = "аеёиоуыэюяАЕЁИОУЫЭЮЯ"
konsonanten = "бвгджзйклмнпрстфхцчшщъьБВГДЖЗЙКЛМНПРСТФХЦЧШЩЪЬ"
grapheme_chars = vokale + konsonanten
satzende = ".!?…"
sonstiges = "„*¤/(|),;:-_\"' “«—»[<>]”–@\r\n\t{}"
allegrapheme = grapheme_chars + satzende + "0123456789" + sonstiges

klammern_auslassen = True
opt_tg = False

# ==================================
def klammern_weg(text):
    pattern = r'\[[^\[\]]*\]'
    while re.search(pattern, text):
        text = re.sub(pattern, '', text)
    return text

def text_vorbereiten(text):
    text = text.replace("\n", " ")
    if klammern_auslassen:
        text = klammern_weg(text)
    return text

def finde_worte(satz):
    if opt_tg:
        pattern = r'{[^}]*}|[' + re.escape(grapheme_chars) + r']+'
    else:
        pattern = r'[' + re.escape(grapheme_chars) + r']+'
    return re.findall(pattern, satz)

def finde_saetze(text):
    text = re.sub(r'\.{2,}', '.', text)
    text = text.replace('…', '.')
    upper_cyr = "А-ЯЁ"
    pattern = (
        r'([' + re.escape(allegrapheme) + r']+?[.!?])'
        r'(?=[^' + re.escape(grapheme_chars) + r']*(?:[' + upper_cyr + r']|$))'
    )
    return [m.group(1) for m in re.finditer(pattern, text)]

def zaehle_silben(wort):
    return sum(1 for c in wort if c in vokale)

def baue_graphem_und_silbenstruktur(text, nullsilber=False):
    saetze = finde_saetze(text)
    grapheme_liste = []
    silben_liste = []

    for satz in saetze:
        worte = finde_worte(satz)
        grapheme_satz = []
        silben_satz = []
        for wort in worte:
            s = zaehle_silben(wort)
            g = len(wort)
            if not nullsilber and s == 0:
                continue
            grapheme_satz.append(g)
            silben_satz.append(s)
        if grapheme_satz:
            grapheme_liste.append(grapheme_satz)
            silben_liste.append(silben_satz)
    return grapheme_liste, silben_liste

# ==================================
def berechne_statistik(text, nullsilber=False):
    text = text_vorbereiten(text)
    grapheme, silben = baue_graphem_und_silbenstruktur(text, nullsilber)

    wortanzahl_text = sum(len(satz) for satz in grapheme)
    satzanzahl_text = len(grapheme)
    silbenanzahl_text = sum(sum(satz) for satz in silben)
    graphemanzahl_text = sum(sum(satz) for satz in grapheme)

    ASL = wortanzahl_text / satzanzahl_text if satzanzahl_text else 0
    AWL = silbenanzahl_text / wortanzahl_text if wortanzahl_text else 0

    # Flesch
    flesch = 206.835 - (84.6 * silbenanzahl_text/wortanzahl_text) - (1.015 * wortanzahl_text/satzanzahl_text) if wortanzahl_text and satzanzahl_text else 0
    # Amstad
    amstad = 180 - ASL - AWL*58.5
    # Tuldava
    tuldava = AWL * math.log(ASL) if ASL>0 else 0
    # Lix
    lange_worte = sum(1 for satz in grapheme for g in satz if g >= 7)
    lix = ASL + (lange_worte / wortanzahl_text * 100) if satzanzahl_text and wortanzahl_text else 0

    # Wiener Sachtextformel (WSTF)
    MS = sum(1 for satz in silben for s in satz if s>=3)/wortanzahl_text*100 if wortanzahl_text else 0
    IW = sum(1 for satz in grapheme for g in satz if g>6)/wortanzahl_text*100 if wortanzahl_text else 0
    ES = sum(1 for satz in silben for s in satz if s==1)/wortanzahl_text*100 if wortanzahl_text else 0
    SL = ASL
    WSTF1 = 0.1935*MS + 0.1672*SL + 0.1297*IW - 0.0327*ES -0.875
    WSTF2 = 0.2007*MS + 0.1682*SL + 0.1373*IW - 2.779
    WSTF3 = 0.2963*MS + 0.1905*SL - 1.1144
    WSTF4 = 0.2656*SL + 0.2744*MS - 1.693

    # New Reading Ease (NRE)
    NOSW = sum(1 for satz in silben for s in satz if s==1)/wortanzahl_text*100 if wortanzahl_text else 0
    NRE = 1.599*NOSW - 1.015*SL -31.517

    # === Gunning-Fog Index ===
    # silben: Liste von Sätzen, jeder Satz ist Liste der Silben pro Wort
    # grapheme: Liste von Sätzen, jeder Satz ist Liste der Graphemanzahl pro Wort

    wortanzahl_text = sum(len(satz) for satz in grapheme)
    asl = wortanzahl_text / len(grapheme) if grapheme else 0

    # Lange Wörter: Wörter mit 3 oder mehr Silben
    lange_worte_gf = sum(1 for satz in silben for w in satz if w >= 3)

    gunning_fog = 0.4 * (asl + 100 * (lange_worte_gf / wortanzahl_text)) if wortanzahl_text else 0.0

    return {
        "Sätze": satzanzahl_text,
        "Wörter": wortanzahl_text,
        "Silben": silbenanzahl_text,
        "Grapheme": graphemanzahl_text,
        "ASL": round(ASL,2),
        "AWL": round(AWL,2),
        "Flesch": round(flesch,2),
        "Amstad": round(amstad,2),
        "Tuldava": round(tuldava,2),
        "Lix": round(lix,2),
        "WSTF1": round(WSTF1,2),
        "WSTF2": round(WSTF2,2),
        "WSTF3": round(WSTF3,2),
        "WSTF4": round(WSTF4,2),
        "NRE": round(NRE,2),
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


# ==================================
class MainGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CYIW ⋅ Calculate Your Index Well ⋅ Russian 1.5")
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
        ToolTip(b2, "Als TXT speichern")
        
        btn_reset = tk.Button(button_frame, text="♻️", command=self.reset_ausgabe, font=("Arial", 20), width=2, height=1)
        btn_reset.pack(side='left', padx=5)
        ToolTip(btn_reset, "Zurücksetzen")

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

  # ===========================
    # Korrelationsanzeige (wie in PSPP)
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

        # Pearson-Korrelation und p-Werte berechnen
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


# ===========================

if __name__=="__main__":
    root = tk.Tk()
    app = MainGUI(root)
    root.mainloop()
