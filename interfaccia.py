#!/usr/bin/env python3
"""
Interfaccia grafica dello scraper dei Manifesti degli Studi PoliMi.
Avvio: avvia.bat (Windows), ./avvia.sh (Linux/macOS) oppure  python avvia.py

La finestra ha tre schede: le prime due sono i due passi d'uso, la terza le ricerche al volo.
    SchedaScarica  scelte dell'utente -> scarica_manifesti.Opzioni -> scarica_manifesti.scarica()
    SchedaEsplora  file JSON -> esporta.tabelle() -> filtri -> esporta.esporta*()
    SchedaCerca    campi della ricerca -> cerca.<ricerca>() -> cerca.Risultato -> tabella, salvataggio
Questo file contiene solo l'interfaccia: la logica sta in scarica_manifesti.py, esporta.py e cerca.py.

Lo scaricamento gira in un thread separato, così la finestra resta reattiva. Il thread non tocca
mai i widget (tkinter non lo permette): manda messaggi nella coda self.q, che il thread della
finestra legge ogni 100 ms (_svuota_coda) e trasforma in aggiornamenti di registro e barra.
"""
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import webbrowser
from datetime import date, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from tkinter.scrolledtext import ScrolledText

import cerca
import esporta as ex
import scarica_manifesti as ps

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "output"
CACHE = HERE / "cache"
VUOTO = "— scegli —"             # combobox in cui l'utente non ha ancora scelto
TUTTI_TIPI = "Tutti i tipi"       # «Tipo di laurea»: nessun filtro
TUTTI = "(tutti)"                 # filtri della scheda 2: nessun filtro
ON, OFF, MEZZO = "☑", "☐", "◩"    # spunta di un corso / gruppo tutto, per niente o in parte selezionato
MAX_ANTEPRIMA = 3000              # righe mostrate nell'anteprima (il salvataggio le include tutte)
GRIGIO, ROSSO, VERDE = "#666", "#b00020", "#1b7f3b"

COSA_FA_BREVE = ("Scarica dai Manifesti degli Studi PoliMi insegnamenti, scaglioni e orari, "
                 "e li salva in Excel, CSV, pagina web o come orario da stampare. "
                 "Oppure cerca al volo chi insegna cosa, quando e dove.")
COSA_FA = ("Scarica dal sito dei Manifesti degli Studi del Politecnico di Milano gli insegnamenti dei corsi "
           "che scegli, con gli scaglioni (gli studenti divisi per iniziale del cognome, ognuno con i suoi "
           "docenti) e l'orario delle lezioni. Poi puoi filtrarli e salvarli in Excel, CSV, pagina web "
           "o come orario settimanale da stampare.")

# Testo della finestra «Guida»: (titolo, paragrafo)
GUIDA = [
    ("Cosa fa questo programma", COSA_FA + "\n\nI dati vengono letti dal sito ufficiale, pubblico: il programma "
     "non modifica nulla e non richiede credenziali."),
    ("Come si usa, in due passi",
     "1 · Scheda «Scarica dati dal sito»: scegli anno accademico, sede, i corsi e cosa includere, poi premi "
     "«Avvia scaricamento». Prima di partire vedi un riepilogo delle scelte da confermare. Il risultato è "
     "un file di dati nella cartella output.\n"
     "2 · Scheda «Esplora ed esporta»: apri un file di dati, scegli una tabella, restringi i risultati con "
     "i filtri e salva le righe che vedi.\n\n"
     "Se hai già scaricato i dati in passato puoi andare direttamente alla scheda 2."),
    ("Scheda 3 · Cerca sul sito",
     "Per una domanda precisa non serve scaricare nulla: la scheda «Cerca sul sito» interroga il sito al volo "
     "e risponde in pochi secondi.\n"
     "• Chi insegna un insegnamento: i docenti, ognuno con il suo orario e il suo scaglione. Con le fasce "
     "orarie (es. «gio 08:15-10:15, ven 10:15-13:15») in cima trovi, evidenziato in verde, chi le copre tutte. "
     "È il modo per scoprire il docente quando conosci solo l'orario delle lezioni.\n"
     "• Scheda di un docente: insegnamenti, scaglioni, orario e contatti.\n"
     "• Insegnamenti e docenti: chi insegna cosa, cercando per insegnamento o per docente.\n"
     "• Aule: chi occupa ogni aula in un giorno, oppure quali aule sono libere in una fascia oraria.\n"
     "• Informazioni su un corso: piani di studio, docenti, programmi interdisciplinari, scambi.\n"
     "• Vecchi ordinamenti: insegnamenti precedenti al D.M. 509.\n\n"
     "Doppio clic su una riga per vederla per intero; da lì (o con il pulsante sotto la tabella) apri la "
     "scheda del docente o la pagina sul sito. I risultati si salvano in Excel, CSV, pagina web o JSON.\n\n"
     "Fasce orarie: un intervallo (gio 08:15-10:15) deve essere coperto tutto da una lezione; un'ora sola "
     "(gio 08:15) vuol dire «a lezione in quel momento». Si possono scrivere a mano oppure comporre con "
     "giorno, «dalle», «alle» e «＋ Aggiungi fascia»."),
    ("Parole da conoscere",
     "• Corso di studi: per esempio Ingegneria Informatica. Ogni corso ha un codice numerico.\n"
     "• Tipo di laurea: Laurea (primo livello), Laurea Magistrale, Ciclo Unico… «ord. 96/23» indica il "
     "regolamento (ordinamento) del corso.\n"
     "• Anno di corso: 1°, 2°, 3°… anno.\n"
     "• Piano di studi: una variante dello stesso corso (per esempio in italiano o in inglese, o in "
     "un'altra sede). Ogni piano ha il suo elenco di insegnamenti. Il piano «***» raccoglie gli "
     "insegnamenti comuni a tutti i piani.\n"
     "• Periodo didattico: 1° semestre, 2° semestre o annuale. Alcuni corsi sono divisi in trimestri: "
     "rientrano in «Altri periodi».\n"
     "• Scaglione: quando un insegnamento ha molti studenti, li divide per iniziale del cognome; ogni "
     "gruppo ha i suoi docenti, orari e aule. «BRU – CON» vuol dire: cognomi da BRU (compreso) fino a "
     "CON (escluso). «A – ZZZZ (unico)» vuol dire che c'è un solo gruppo per tutti.\n"
     "• Fascia di cognomi: gli insegnamenti dividono i cognomi in modi diversi. La tabella «Orario per "
     "cognome» calcola le fasce in cui la divisione è la stessa per tutti gli insegnamenti, e per ognuna "
     "mostra l'orario completo di uno studente con quel cognome."),
    ("Le tabelle della scheda 2", "\n".join(f"• {nome}: {ex.DESCRIZIONI[k]}." for k, nome in [
        ("insegnamenti", "Insegnamenti"), ("scaglioni", "Scaglioni"), ("cognomi", "Orario per cognome"),
        ("lezioni", "Lezioni (orario)"), ("piani", "Corsi e piani")])
     + "\n\nLo stesso insegnamento può comparire in più corsi e piani, e quindi in più righe: l'opzione "
       "«Una riga sola per ciò che si ripete in più corsi» le unisce."),
    ("Cache: le pagine già scaricate",
     "Ogni pagina letta dal sito resta salvata nella cartella cache, così ripetere o ampliare uno "
     "scaricamento è immediato. Se il sito è stato aggiornato e vuoi i dati nuovi, premi «Svuota cache» "
     "prima di scaricare."),
    ("Se qualcosa va storto",
     "• Puoi interrompere in qualunque momento (anche chiudendo la finestra): i dati raccolti fino a quel "
     "punto vengono salvati comunque. Lo stesso vale se la connessione cade.\n"
     "• Se un insegnamento non si riesce a leggere, gli altri continuano: lo trovi segnalato nella colonna "
     "«Note». Rilanciando lo scaricamento, le pagine già lette vengono prese dalla cache.\n"
     "• «Non riesco a scrivere il file»: probabilmente è aperto in Excel. Chiudilo e riprova."),
]


def tipo_laurea(gruppo):
    """'Laurea Magistrale - ord. 96/23' -> 'Laurea Magistrale' (il tipo senza l'ordinamento)."""
    return re.split(r"\s+-\s+ord\b", gruppo or "", maxsplit=1, flags=re.I)[0].strip() or "Altro"


def apri_percorso(path):
    """Apre un file o una cartella con il programma predefinito del sistema."""
    path = str(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Impossibile aprire", f"{path}\n\n{e}")


class Sezione(ttk.LabelFrame):
    """Riquadro con titolo che raggruppa i controlli di un passo (① ② ③ …)."""
    def __init__(self, parent, titolo, **kw):
        super().__init__(parent, text=f" {titolo} ", padding=(10, 6), **kw)


class Scorrevole(ttk.Frame):
    """Contenitore con barra di scorrimento verticale, che compare solo se il contenuto non ci sta.
    I controlli vanno messi dentro self.interno."""

    def __init__(self, parent):
        super().__init__(parent)
        sfondo = ttk.Style(self).lookup("TFrame", "background") or None
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=sfondo)
        self.barra = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.barra.set)
        self.interno = ttk.Frame(self.canvas)
        self._finestra = self.canvas.create_window(0, 0, window=self.interno, anchor="nw")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.interno.bind("<Configure>", self._adatta)
        self.canvas.bind("<Configure>", self._adatta)
        # rotella del mouse: attiva solo mentre il puntatore è sopra il contenitore
        self.bind("<Enter>", lambda e: self.bind_all("<MouseWheel>", self._rotella) or
                  self.bind_all("<Button-4>", self._rotella) or self.bind_all("<Button-5>", self._rotella))
        self.bind("<Leave>", lambda e: [self.unbind_all(s) for s in ("<MouseWheel>", "<Button-4>", "<Button-5>")])

    def _adatta(self, _event=None):
        # il canvas chiede la dimensione del contenuto, così su schermi grandi non serve scorrere
        self.canvas.configure(width=self.interno.winfo_reqwidth(), height=self.interno.winfo_reqheight(),
                              scrollregion=(0, 0, 0, self.interno.winfo_reqheight()))
        self.canvas.itemconfigure(self._finestra, width=self.canvas.winfo_width())
        if self.interno.winfo_reqheight() > self.canvas.winfo_height() > 1:
            self.barra.grid(row=0, column=1, sticky="ns")
        else:
            self.barra.grid_remove()
            self.canvas.yview_moveto(0)

    def _rotella(self, event):
        if self.barra.winfo_ismapped():
            passi = -1 if getattr(event, "num", None) == 4 or event.delta > 0 else 1
            self.canvas.yview_scroll(passi, "units")


class Suggerimento:
    """Spiegazione che compare tenendo il mouse fermo su un controllo per mezzo secondo."""

    def __init__(self, widget, testo, attesa_ms=500):
        self.widget, self.testo, self.attesa_ms = widget, testo, attesa_ms
        self._timer = self._finestra = None
        widget.bind("<Enter>", self._programma, add="+")
        widget.bind("<Leave>", self._nascondi, add="+")
        widget.bind("<ButtonPress>", self._nascondi, add="+")

    def _programma(self, _event=None):
        self._nascondi()
        self._timer = self.widget.after(self.attesa_ms, self._mostra)

    def _mostra(self):
        self._timer = None
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._finestra = tk.Toplevel(self.widget)
        self._finestra.wm_overrideredirect(True)  # senza bordo né barra del titolo
        self._finestra.wm_geometry(f"+{x}+{y}")
        tk.Label(self._finestra, text=self.testo, justify="left", wraplength=440, padx=8, pady=5,
                 background="#ffffe6", foreground="#1c1e21", relief="solid", borderwidth=1).pack()

    def _nascondi(self, _event=None):
        if self._timer:
            self.widget.after_cancel(self._timer)
            self._timer = None
        if self._finestra:
            self._finestra.destroy()
            self._finestra = None


def salva_file(widget, path, scrivi, testo):
    """Esegue scrivi(); se va bene propone di aprire il file, altrimenti spiega il problema."""
    try:
        widget.config(cursor="watch")
        widget.update_idletasks()
        scrivi()
    except PermissionError:
        messagebox.showerror("File in uso", "Non riesco a scrivere il file: forse è aperto in Excel "
                             "o in un altro programma? Chiudilo e riprova.")
        return
    except Exception as e:
        messagebox.showerror("Salvataggio non riuscito", f"{path}\n\n{e}")
        return
    finally:
        widget.config(cursor="")
    if messagebox.askyesno("Salvato", f"{testo}\n\n{path}\n\nVuoi aprirlo adesso?"):
        apri_percorso(path)


def aiuto(testo, *widgets):
    """Stessa spiegazione su più controlli (di solito l'etichetta e il campo accanto)."""
    for w in widgets:
        Suggerimento(w, testo)


def mostra_guida(parent):
    """Finestra con la guida all'uso e il significato dei termini (testo in GUIDA)."""
    w = tk.Toplevel(parent)
    w.title("Guida")
    w.geometry("760x620")
    t = ScrolledText(w, wrap="word", font=("", 10), padx=14, pady=10, spacing2=2)
    t.pack(fill="both", expand=True)
    t.tag_config("titolo", font=("", 12, "bold"), spacing1=10, spacing3=4)
    for titolo, testo in GUIDA:
        t.insert("end", titolo + "\n", "titolo")
        t.insert("end", testo + "\n")
    t.config(state="disabled")
    ttk.Button(w, text="Chiudi", command=w.destroy).pack(pady=6)


# ============================================================ scheda 1: scarica

class SchedaScarica(ttk.Frame):
    def __init__(self, app, parent):
        super().__init__(parent, padding=10)
        self.app = app
        self.catalogo = None
        self.map_aa, self.map_sede = {}, {}
        self.n_richiesta_corsi = 0    # solo l'ultima richiesta del catalogo conta
        self.selezionati = set()      # codici corso selezionati
        self.q = queue.Queue()
        self.stop = threading.Event()
        self.lavoro = None
        self._fase, self._t0, self._n0 = None, 0.0, 0  # per la stima del tempo rimanente
        self._costruisci()
        self._carica_opzioni_iniziali()
        self.after(100, self._svuota_coda)

    # ---------------------------------------------------------------- layout
    def _costruisci(self):
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=2)

        # ① anno e sede
        s1 = Sezione(self, "① Anno accademico e sede")
        s1.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        lbl = ttk.Label(s1, text="Anno accademico")
        lbl.grid(row=0, column=0, sticky="w")
        self.cb_aa = ttk.Combobox(s1, state="readonly", width=14, values=[VUOTO])
        self.cb_aa.grid(row=0, column=1, padx=(6, 18))
        aiuto("L'anno accademico dei manifesti da leggere. L'elenco arriva direttamente dal sito.", lbl, self.cb_aa)
        lbl = ttk.Label(s1, text="Sede")
        lbl.grid(row=0, column=2, sticky="w")
        self.cb_sede = ttk.Combobox(s1, state="readonly", width=26, values=[VUOTO])
        self.cb_sede.grid(row=0, column=3, padx=(6, 18))
        aiuto("La sede (città) dei corsi. Scegliendo una sede vengono scaricati solo i piani di studio "
              "erogati lì, salvo diversa scelta in «Cosa includere».", lbl, self.cb_sede)
        lbl = ttk.Label(s1, text="Tipo di laurea")
        lbl.grid(row=0, column=4, sticky="w")
        self.cb_tipo = ttk.Combobox(s1, state="readonly", width=30, values=[TUTTI_TIPI])
        self.cb_tipo.set(TUTTI_TIPI)
        self.cb_tipo.grid(row=0, column=5, padx=(6, 18))
        self.cb_tipo.bind("<<ComboboxSelected>>", lambda e: self._riempi_albero())
        aiuto("Mostra solo i corsi di un tipo (es. Laurea Magistrale). Vengono scaricati solo i corsi "
              "selezionati di questo tipo: quelli spuntati ma di un altro tipo restano esclusi.", lbl, self.cb_tipo)
        self.lbl_stato = ttk.Label(s1, text="Collegamento al sito…", foreground=GRIGIO)
        self.lbl_stato.grid(row=0, column=6, sticky="w")
        for cb in (self.cb_aa, self.cb_sede):
            cb.bind("<<ComboboxSelected>>", lambda e: self._carica_corsi())

        # ② corsi
        s2 = Sezione(self, "② Corsi di studio (clic per selezionare; clic su un gruppo = tutto il gruppo)")
        s2.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        s2.columnconfigure(1, weight=1)
        s2.rowconfigure(1, weight=1)
        lbl = ttk.Label(s2, text="Cerca")
        lbl.grid(row=0, column=0, sticky="w")
        self.var_cerca = tk.StringVar()
        self.var_cerca.trace_add("write", lambda *a: self._riempi_albero())
        e = ttk.Entry(s2, textvariable=self.var_cerca)
        e.grid(row=0, column=1, sticky="ew", padx=6)
        aiuto("Filtra l'elenco per nome, codice o scuola (es. «informatica»). Serve solo a trovare i corsi: "
              "quelli già selezionati restano selezionati anche se non si vedono.", lbl, e)
        b = ttk.Button(s2, text="Seleziona visibili", command=self._seleziona_visibili)
        b.grid(row=0, column=2)
        aiuto("Seleziona tutti i corsi che si vedono ora nell'elenco (per esempio dopo una ricerca).", b)
        ttk.Button(s2, text="Deseleziona tutti", command=self._deseleziona_tutti).grid(row=0, column=3, padx=(4, 0))
        fr = ttk.Frame(s2)
        fr.grid(row=1, column=0, columnspan=4, sticky="nsew", pady=6)
        fr.columnconfigure(0, weight=1)
        fr.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(fr, columns=("codice",), show="tree headings", selectmode="none")
        self.tree.heading("#0", text="Scuola / tipo di laurea / corso")
        self.tree.heading("codice", text="Codice")
        self.tree.column("#0", width=520, stretch=True)
        self.tree.column("codice", width=70, anchor="center", stretch=False)
        sb = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<Button-1>", self._clic_albero)
        self.tree.bind("<space>", self._clic_albero)
        self.lbl_sel = ttk.Label(s2, text="Scegli anno accademico e sede per vedere i corsi.")
        self.lbl_sel.grid(row=2, column=0, columnspan=4, sticky="w")

        # ③ opzioni
        riquadro = Sezione(self, "③ Cosa includere")
        riquadro.grid(row=1, column=1, sticky="nsew")
        riquadro.columnconfigure(0, weight=1)
        riquadro.rowconfigure(0, weight=1)
        scorrevole = Scorrevole(riquadro)  # su schermi bassi le opzioni non ci stanno tutte
        scorrevole.grid(row=0, column=0, sticky="nsew")
        s3 = scorrevole.interno
        s3.columnconfigure(0, weight=1)
        r = 0

        def titolo(testo, spiegazione):
            nonlocal r
            lbl = ttk.Label(s3, text=testo, font=("", 9, "bold"))
            lbl.grid(row=r, column=0, sticky="w", pady=(8 if r else 0, 0))
            aiuto(spiegazione, lbl)
            r += 1

        def casella(parent, testo, var, spiegazione, **grid):
            cb = ttk.Checkbutton(parent, text=testo, variable=var, command=grid.pop("command", None))
            cb.grid(**grid)
            aiuto(spiegazione, cb)

        titolo("Anni di corso", "Di quali anni di corso scaricare gli insegnamenti.")
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        self.var_anni = {}
        for i, a in enumerate(["1", "2", "3", "4", "5", "6"]):
            v = tk.BooleanVar()
            casella(f, f"{a}°", v, f"Gli insegnamenti del {a}° anno. Se un corso non ha questo anno "
                    "(es. 4° anno di una laurea triennale) viene saltato.",
                    command=self._anni_cambiati, row=0, column=i, padx=(0, 6))
            self.var_anni[a] = v
        self.var_anni_tutti = tk.BooleanVar()
        casella(f, "Tutti insieme", self.var_anni_tutti,
                "Tutti gli anni di corso con una sola lettura per piano: è più veloce che spuntarli uno per uno. "
                "L'anno di ogni insegnamento resta indicato nella colonna «Anno».",
                command=self._anni_tutti_cambiato, row=0, column=6, padx=(6, 0))

        titolo("Periodo didattico", "Tiene solo gli insegnamenti che si svolgono nei periodi spuntati.")
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        self.var_periodi = {}
        spiegazioni = {"annuale": "Insegnamenti che durano tutto l'anno (entrambi i semestri).",
                       "1sem": "Insegnamenti del primo semestre (circa settembre–gennaio).",
                       "2sem": "Insegnamenti del secondo semestre (circa febbraio–luglio).",
                       "altro": "Periodi diversi dai precedenti, se il sito ne indica: per esempio i trimestri "
                                "di alcuni corsi (come Industrial Engineering a Piacenza) o i corsi brevi."}
        for i, (k, t) in enumerate(ps.PERIODI.items()):
            v = tk.BooleanVar()
            casella(f, t, v, spiegazioni[k], row=0, column=i, padx=(0, 8))
            self.var_periodi[k] = v

        titolo("Piani di studio", "Un corso può avere più piani di studio: varianti dello stesso corso "
               "(es. in italiano o in inglese, o in un'altra sede), ognuna con i suoi insegnamenti.")
        self.var_piani = tk.StringVar(value="")
        for valore, testo, spiegazione in [
                ("tutti", "Tutti i piani della sede", "Scarica tutti i piani di studio di ogni corso scelto."),
                ("primo", "Solo il primo piano di ogni corso",
                 "Un solo piano per corso (il primo dell'elenco del sito): più veloce, utile per una prima occhiata.")]:
            rb = ttk.Radiobutton(s3, text=testo, value=valore, variable=self.var_piani)
            rb.grid(row=r, column=0, sticky="w"); r += 1
            aiuto(spiegazione, rb)
        self.var_altre_sedi = tk.BooleanVar()
        casella(s3, "Tieni anche i piani erogati in altre sedi (es. Cremona)", self.var_altre_sedi,
                "Alcuni corsi hanno piani in più città. Normalmente si tengono solo quelli della sede scelta in "
                "alto; con questa opzione si tengono tutti. I piani scartati sono comunque elencati nella tabella "
                "«Corsi e piani».", row=r, column=0, sticky="w"); r += 1
        self.var_nondiv = tk.BooleanVar()
        casella(s3, "Includi il piano «***» (insegnamenti comuni a tutti i piani)", self.var_nondiv,
                "Sul sito il piano «***» (offerta non diversificata) elenca gli insegnamenti non legati a un "
                "piano specifico, uguali per tutti. Normalmente non viene scaricato: spunta per includerlo.",
                row=r, column=0, sticky="w"); r += 1

        titolo("Dettagli da scaricare", "Oltre all'elenco degli insegnamenti, cosa leggere dalla pagina di "
               "ognuno. Ogni insegnamento richiede una o due pagine in più: è la parte che richiede più tempo.")
        self.var_scaglioni = tk.BooleanVar()
        self.var_orari = tk.BooleanVar()
        casella(s3, "Scaglioni (lettere, docenti, moduli)", self.var_scaglioni,
                "Per ogni insegnamento, i gruppi di studenti divisi per iniziale del cognome, con i docenti "
                "di ciascun gruppo.", command=self._dettagli_cambiati, row=r, column=0, sticky="w"); r += 1
        casella(s3, "Orario delle lezioni (giorni, ore, aule)", self.var_orari,
                "L'orario settimanale di ogni scaglione: giorno, ora, aula, date. Richiede gli scaglioni, "
                "che vengono spuntati da sé.", command=self._orari_cambiati, row=r, column=0, sticky="w"); r += 1
        ttk.Label(s3, text="Senza dettagli si ottiene solo l'elenco degli insegnamenti (molto più veloce).",
                  foreground=GRIGIO, wraplength=520).grid(row=r, column=0, sticky="w"); r += 1

        # impostazioni che servono di rado: in una finestra a parte (vedi _avanzate)
        self.var_cache = tk.BooleanVar(value=True)
        self.var_delay = tk.DoubleVar(value=0.4)
        self.var_paralleli = tk.IntVar(value=ps.PARALLELI_DEFAULT)
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w", pady=(10, 0)); r += 1
        b = ttk.Button(f, text="Impostazioni avanzate…", command=self._avanzate)
        b.grid(row=0, column=0)
        aiuto("Cache delle pagine già scaricate, pausa tra le richieste e richieste in parallelo. "
              "Di solito non serve cambiarle.", b)
        self.lbl_cache = ttk.Label(f, foreground=GRIGIO)
        self.lbl_cache.grid(row=0, column=1, padx=8)
        self._aggiorna_info_cache()

        # ④ avvio
        s4 = Sezione(self, "④ Scarica")
        s4.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        s4.columnconfigure(1, weight=1)
        lbl = ttk.Label(s4, text="Salva in")
        lbl.grid(row=0, column=0, sticky="w")
        self.var_out = tk.StringVar()
        e = ttk.Entry(s4, textvariable=self.var_out)
        e.grid(row=0, column=1, sticky="ew", padx=6)
        aiuto("Il file di dati (.json) in cui salvare il risultato. Se lo lasci vuoto viene creato un file "
              "nuovo nella cartella output, con anno, sede, data e ora nel nome.", lbl, e)
        ttk.Button(s4, text="Sfoglia…", command=self._scegli_out).grid(row=0, column=2)
        ttk.Label(s4, text="(vuoto = nome automatico nella cartella output)", foreground=GRIGIO).grid(row=0, column=3, padx=6)
        f = ttk.Frame(s4); f.grid(row=1, column=0, columnspan=4, sticky="ew", pady=6)
        f.columnconfigure(3, weight=1)
        self.btn_avvia = ttk.Button(f, text="▶  Avvia scaricamento", style="Accent.TButton", command=self._avvia)
        self.btn_avvia.grid(row=0, column=0, ipadx=10, ipady=3)
        aiuto("Controlla le scelte, mostra un riepilogo da confermare e avvia lo scaricamento.", self.btn_avvia)
        self.btn_stop = ttk.Button(f, text="■  Interrompi", command=self._interrompi, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=8)
        aiuto("Ferma lo scaricamento. I dati raccolti fino a questo momento vengono salvati comunque.", self.btn_stop)
        self.lbl_prog = ttk.Label(f, text="", width=58)
        self.lbl_prog.grid(row=0, column=2, padx=8)
        self.pbar = ttk.Progressbar(f, mode="determinate")
        self.pbar.grid(row=0, column=3, sticky="ew")
        self.log = ScrolledText(s4, height=6, font=("Consolas" if sys.platform.startswith("win") else "Monospace", 9),
                                state="disabled", wrap="none")
        self.log.grid(row=2, column=0, columnspan=4, sticky="nsew")
        s4.rowconfigure(2, weight=1)

        # lo spazio in più (o in meno) va soprattutto a corsi e opzioni, meno al registro
        self.rowconfigure(1, weight=3)
        self.rowconfigure(2, weight=1)

    # ---------------------------------------------------------------- catalogo
    def _in_thread(self, funzione, fine):
        """Esegue funzione() in un thread; poi, nel thread della finestra, fine(risultato, errore)."""
        def run():
            try:
                self.q.put(("fine", fine, funzione(), None))
            except Exception as e:
                self.q.put(("fine", fine, None, e))
        threading.Thread(target=run, daemon=True).start()

    def _carica_opzioni_iniziali(self):
        def f():
            home = ps.soup(ps.Client(delay=0).get(ps.BASE, {"evn_DEFAULT": "evento", "lang": "IT"}))
            return ps.select_options(home, "aa"), ps.select_options(home, "sede")
        self._in_thread(f, self._opzioni_caricate)

    def _opzioni_caricate(self, res, err):
        if err:
            self.lbl_stato.config(text="Sito non raggiungibile: controlla la connessione", foreground=ROSSO)
            messagebox.showerror("Errore di connessione", f"Impossibile raggiungere il sito dei manifesti.\n\n{err}")
            return
        anni, sedi = res
        self.map_aa = {o["testo"]: o["valore"] for o in anni}
        self.map_sede = {o["testo"]: o["valore"] for o in sedi}
        self.cb_aa.config(values=list(self.map_aa))
        self.cb_sede.config(values=list(self.map_sede))
        self.cb_aa.set(VUOTO)
        self.cb_sede.set(VUOTO)
        self.lbl_stato.config(text="Scegli anno accademico e sede", foreground=GRIGIO)

    def _carica_corsi(self):
        aa, sede = self.map_aa.get(self.cb_aa.get()), self.map_sede.get(self.cb_sede.get())
        if not (aa and sede):
            return
        self.lbl_stato.config(text="Carico l'elenco dei corsi…", foreground=GRIGIO)
        self.catalogo = None
        self._riempi_albero()
        self.n_richiesta_corsi += 1
        n = self.n_richiesta_corsi
        self._in_thread(lambda: ps.catalogo(aa, sede), lambda cat, err: self._corsi_caricati(cat, err, n))

    def _corsi_caricati(self, cat, err, n_richiesta):
        if n_richiesta != self.n_richiesta_corsi:
            return  # nel frattempo l'utente ha cambiato anno o sede
        if err:
            self.lbl_stato.config(text="Errore nel caricare i corsi", foreground=ROSSO)
            messagebox.showerror("Errore", str(err))
            return
        self.catalogo = cat
        tipi = sorted({tipo_laurea(c["gruppo"]) for s in cat["scuole"] for c in s["corsi"]})
        self.cb_tipo.config(values=[TUTTI_TIPI] + tipi)
        if self.cb_tipo.get() not in tipi:
            self.cb_tipo.set(TUTTI_TIPI)
        validi = {c["codice"] for s in cat["scuole"] for c in s["corsi"]}
        self.selezionati &= validi
        n = len(validi)
        self.lbl_stato.config(text=f"{n} corsi disponibili", foreground=VERDE)
        self._riempi_albero()

    # ---------------------------------------------------------------- albero
    def _selezione_effettiva(self):
        """Corsi selezionati che rispettano il «Tipo di laurea» scelto: sono quelli che verranno scaricati.
        (La casella «Cerca» invece non esclude nulla: serve solo a trovare i corsi.)"""
        if not self.catalogo:
            return set()
        tipo = self.cb_tipo.get()
        return {c["codice"] for s in self.catalogo["scuole"] for c in s["corsi"]
                if c["codice"] in self.selezionati and (tipo == TUTTI_TIPI or tipo_laurea(c["gruppo"]) == tipo)}

    def _corsi_visibili(self):
        if not self.catalogo:
            return []
        parole = self.var_cerca.get().lower().split()
        tipo = self.cb_tipo.get()
        out = []
        for s in self.catalogo["scuole"]:
            for c in s["corsi"]:
                if tipo != TUTTI_TIPI and tipo_laurea(c["gruppo"]) != tipo:
                    continue
                testo = f"{s['nome']} {c['gruppo']} {c['nome']} {c['codice']}".lower()
                if all(p in testo for p in parole):
                    out.append((s, c))
        return out

    def _riempi_albero(self):
        aperti = {self.tree.item(i, "text")[2:] for i in self._tutti_nodi() if self.tree.item(i, "open")}
        self.tree.delete(*self.tree.get_children())
        self.nodi = {}  # iid -> set di codici corso sotto il nodo
        filtro = bool(self.var_cerca.get().strip()) or self.cb_tipo.get() != TUTTI_TIPI
        for s, c in self._corsi_visibili():
            sid = f"s{s['codice']}"
            if not self.tree.exists(sid):
                self.tree.insert("", "end", iid=sid, text=s["nome"], open=filtro or s["nome"] in aperti or not aperti)
                self.nodi[sid] = set()
            gid = f"{sid}|{c['gruppo']}"
            if not self.tree.exists(gid):
                self.tree.insert(sid, "end", iid=gid, text=c["gruppo"] or "Altro", open=True)
                self.nodi[gid] = set()
            nome = c["nome"].replace(f" ({c['codice']})", "")
            self.tree.insert(gid, "end", iid=f"c{c['codice']}|{gid}", text=nome, values=(c["codice"],))
            self.nodi[sid].add(c["codice"])
            self.nodi[gid].add(c["codice"])
        self._aggiorna_spunte()

    def _tutti_nodi(self, parent=""):
        for i in self.tree.get_children(parent):
            yield i
            yield from self._tutti_nodi(i)

    def _codici_nodo(self, iid):
        if iid.startswith("c"):
            return {iid[1:].split("|")[0]}
        return self.nodi.get(iid, set())

    def _aggiorna_spunte(self):
        for iid in self._tutti_nodi():
            cod = self._codici_nodo(iid)
            sel = len(cod & self.selezionati)
            simbolo = ON if cod and sel == len(cod) else (MEZZO if sel else OFF)
            testo = self.tree.item(iid, "text")
            if testo[:1] in (ON, OFF, MEZZO):
                testo = testo[2:]
            self.tree.item(iid, text=f"{simbolo} {testo}")
        n = len(self._selezione_effettiva())
        esclusi = len(self.selezionati) - n
        testo = f"{n} corsi selezionati" if n else "Nessun corso selezionato"
        if esclusi:
            testo += f" (altri {esclusi} spuntati ma di un altro tipo di laurea: non verranno scaricati)"
        self.lbl_sel.config(text=testo)

    def _clic_albero(self, event):
        if event.type == tk.EventType.KeyPress:
            iid = self.tree.focus()
        else:
            if self.tree.identify_region(event.x, event.y) not in ("tree", "cell"):
                return
            if self.tree.identify_element(event.x, event.y).endswith("indicator"):
                return  # clic sulla freccia: apri/chiudi
            iid = self.tree.identify_row(event.y)
        if not iid:
            return
        cod = self._codici_nodo(iid)
        if cod <= self.selezionati:
            self.selezionati -= cod
        else:
            self.selezionati |= cod
        self.tree.focus(iid)
        self._aggiorna_spunte()
        return "break"

    def _seleziona_visibili(self):
        self.selezionati |= {c["codice"] for _, c in self._corsi_visibili()}
        self._aggiorna_spunte()

    def _deseleziona_tutti(self):
        self.selezionati.clear()
        self._aggiorna_spunte()

    # ---------------------------------------------------------------- opzioni
    # «Tutti insieme» e i singoli anni si escludono a vicenda: sono due modi diversi di chiedere
    # gli anni al sito (una pagina con tutti gli anni, oppure una pagina per anno)
    def _anni_cambiati(self):
        if any(v.get() for v in self.var_anni.values()):
            self.var_anni_tutti.set(False)

    def _anni_tutti_cambiato(self):
        if self.var_anni_tutti.get():
            for v in self.var_anni.values():
                v.set(False)

    # l'orario è diviso per scaglione e si legge dalla stessa pagina: «Orario» porta con sé «Scaglioni»
    def _dettagli_cambiati(self):
        if not self.var_scaglioni.get():
            self.var_orari.set(False)

    def _orari_cambiati(self):
        if self.var_orari.get():
            self.var_scaglioni.set(True)

    def _aggiorna_info_cache(self):
        n = len(list(CACHE.glob("*.html"))) if CACHE.exists() else 0
        mb = sum(f.stat().st_size for f in CACHE.glob("*.html")) / 1e6 if n else 0
        self.lbl_cache.config(text=f"Cache: {n} pagine già scaricate ({mb:.0f} MB)")

    def _avanzate(self):
        """Finestra con le impostazioni che servono di rado. Le variabili (var_cache, var_delay,
        var_paralleli) appartengono alla scheda, così le scelte restano anche chiudendo la finestra."""
        w = tk.Toplevel(self)
        w.title("Impostazioni avanzate")
        w.transient(self.winfo_toplevel())
        fr = ttk.Frame(w, padding=14)
        fr.pack(fill="both", expand=True)
        ttk.Label(fr, text="Di solito non serve cambiare queste impostazioni.", foreground=GRIGIO).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(fr, text="Pagine già scaricate (cache)", font=("", 9, "bold")).grid(row=1, column=0, sticky="w")
        cb = ttk.Checkbutton(fr, text="Riusa le pagine già scaricate", variable=self.var_cache)
        cb.grid(row=2, column=0, columnspan=2, sticky="w")
        aiuto("Le pagine lette dal sito restano salvate (cartella cache): ripetere o ampliare uno scaricamento "
              "diventa immediato. Togli la spunta per rileggere tutto dal sito senza cancellare la cache.", cb)
        b = ttk.Button(fr, text="Svuota cache", command=self._svuota_cache)
        b.grid(row=2, column=2, padx=(12, 0))
        aiuto("Cancella le pagine salvate. Usalo quando vuoi essere sicuro di avere i dati aggiornati dal sito.", b)
        ttk.Label(fr, text="Se il sito è stato aggiornato, svuota la cache per avere i dati nuovi.",
                  foreground=GRIGIO).grid(row=3, column=0, columnspan=3, sticky="w")

        ttk.Label(fr, text="Velocità", font=("", 9, "bold")).grid(row=4, column=0, sticky="w", pady=(12, 0))
        lbl = ttk.Label(fr, text="Pausa tra le richieste (secondi)")
        lbl.grid(row=5, column=0, sticky="w")
        sp = ttk.Spinbox(fr, from_=0.1, to=5, increment=0.1, width=5, textvariable=self.var_delay)
        sp.grid(row=5, column=1, sticky="w", padx=6, pady=2)
        aiuto("Attesa prima di ogni richiesta al sito, per non sovraccaricarlo. Aumentala se il sito "
              "risponde con errori.", lbl, sp)
        lbl = ttk.Label(fr, text="Richieste in parallelo")
        lbl.grid(row=6, column=0, sticky="w")
        sp = ttk.Spinbox(fr, from_=1, to=8, increment=1, width=5, textvariable=self.var_paralleli)
        sp.grid(row=6, column=1, sticky="w", padx=6, pady=2)
        aiuto(f"Quante pagine chiedere al sito contemporaneamente (normale: {ps.PARALLELI_DEFAULT}). "
              "Di più è più veloce ma pesa di più sul sito; 1 = una alla volta.", lbl, sp)

        ttk.Button(fr, text="Chiudi", command=w.destroy).grid(row=7, column=2, sticky="e", pady=(14, 0))
        w.grab_set()

    def _svuota_cache(self):
        if self.in_corso():
            messagebox.showinfo("Svuota cache", "Attendi la fine dello scaricamento in corso.")
            return
        if messagebox.askyesno("Svuota cache", "Cancellare le pagine salvate?\n"
                               "Il prossimo scaricamento rileggerà tutto dal sito (più lento, ma aggiornato)."):
            shutil.rmtree(CACHE, ignore_errors=True)
            self._aggiorna_info_cache()

    def _scegli_out(self):
        OUTPUT.mkdir(exist_ok=True)
        p = filedialog.asksaveasfilename(initialdir=OUTPUT, defaultextension=".json",
                                         filetypes=[("File dati JSON", "*.json")])
        if p:
            self.var_out.set(p)

    def _opzioni(self):
        errori = []
        aa, sede = self.map_aa.get(self.cb_aa.get()), self.map_sede.get(self.cb_sede.get())
        if not aa:
            errori.append("• scegli l'anno accademico")
        if not sede:
            errori.append("• scegli la sede")
        corsi = self._selezione_effettiva()
        if not corsi:
            errori.append("• seleziona almeno un corso di studio (attendi che l'elenco sia caricato)")
        anni = ["0"] if self.var_anni_tutti.get() else [a for a, v in self.var_anni.items() if v.get()]
        if not anni:
            errori.append("• scegli almeno un anno di corso (o «Tutti insieme»)")
        periodi = {k for k, v in self.var_periodi.items() if v.get()}
        if not periodi:
            errori.append("• scegli almeno un periodo didattico")
        if not self.var_piani.get():
            errori.append("• scegli quali piani di studio includere")
        if errori:
            messagebox.showwarning("Mancano delle scelte", "Prima di avviare:\n\n" + "\n".join(errori))
            return None
        try:
            delay = max(0.1, float(self.var_delay.get()))
        except (tk.TclError, ValueError):
            delay = 0.4
        try:
            paralleli = min(8, max(1, int(self.var_paralleli.get())))
        except (tk.TclError, ValueError):
            paralleli = ps.PARALLELI_DEFAULT
        scuole = sorted({s["codice"] for s in self.catalogo["scuole"]
                         for c in s["corsi"] if c["codice"] in corsi})
        return ps.Opzioni(
            aa=aa, sede=sede, scuole=scuole, corsi=sorted(corsi),
            anni_corso=anni, piani=self.var_piani.get(),
            includi_non_diversificato=self.var_nondiv.get(), includi_altre_sedi=self.var_altre_sedi.get(),
            periodi=None if len(periodi) == len(ps.PERIODI) else periodi,
            scaglioni=self.var_scaglioni.get(), orari=self.var_orari.get(),
            out=self.var_out.get().strip() or None, delay=delay, paralleli=paralleli,
            cache=str(CACHE) if self.var_cache.get() else "")

    # ---------------------------------------------------------------- esecuzione
    def _scrivi_log(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        if int(self.log.index("end-1c").split(".")[0]) > 4000:
            self.log.delete("1.0", "500.0")
        self.log.see("end")
        self.log.config(state="disabled")

    def _riepilogo_scelte(self, opt):
        """Le scelte dell'utente in parole, da confermare prima di partire."""
        # con codice, tipo e ordinamento: lo stesso nome può comparire più volte
        # (es. Ingegneria Informatica ord. 96/23 e ord. 270, oppure laurea e magistrale)
        nomi = {c["codice"]: f"{c['nome']} · {c['gruppo']}" for s in self.catalogo["scuole"] for c in s["corsi"]}
        corsi = [nomi.get(c, c) for c in opt.corsi]
        elenco = "\n".join(f"      – {n}" for n in corsi[:8])
        if len(corsi) > 8:
            elenco += f"\n      … e altri {len(corsi) - 8}"
        anni = "tutti" if opt.anni_corso == ["0"] else ", ".join(f"{a}°" for a in opt.anni_corso)
        periodi = "tutti" if opt.periodi is None else ", ".join(t for k, t in ps.PERIODI.items() if k in opt.periodi)
        piani = "tutti i piani" if opt.piani == "tutti" else "solo il primo piano di ogni corso"
        if opt.sede != "ALL_SEDI":
            piani += ", anche di altre sedi" if opt.includi_altre_sedi else ", solo quelli della sede scelta"
        if opt.includi_non_diversificato:
            piani += ", più il piano «***»"
        dettagli = ("scaglioni e orario delle lezioni" if opt.orari else "scaglioni (senza orario)"
                    if opt.scaglioni else "nessuno: solo l'elenco degli insegnamenti")
        cache = ("le pagine già scaricate in passato verranno riusate" if opt.cache
                 else "tutte le pagine verranno rilette dal sito")
        return (f"Stai per scaricare:\n\n"
                f"  • {len(corsi)} {'corso' if len(corsi) == 1 else 'corsi'} di studio:\n{elenco}\n"
                f"  • anno accademico {self.cb_aa.get()}, sede {self.cb_sede.get()}\n"
                f"  • anni di corso: {anni}\n"
                f"  • periodi: {periodi}\n"
                f"  • piani: {piani}\n"
                f"  • dettagli: {dettagli}\n\n"
                f"Nota: {cache}. Puoi interrompere in qualsiasi momento senza perdere i dati raccolti.\n\n"
                f"Avviare lo scaricamento?")

    def _avvia(self):
        opt = self._opzioni()
        if not opt or not messagebox.askyesno("Conferma", self._riepilogo_scelte(opt)):
            return
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self.stop.clear()
        self.btn_avvia.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.pbar.config(value=0, maximum=1)
        self.lbl_prog.config(text="Avvio…")
        self._fase = None

        def run():
            try:
                out = ps.scarica(opt, log=lambda m: self.q.put(("log", m)),
                                 progress=lambda f, n, t: self.q.put(("prog", f, n, t)), stop=self.stop)
                self.q.put(("finito", out, None))
            except Exception as e:
                self.q.put(("finito", None, e))
        self.lavoro = threading.Thread(target=run, daemon=True)
        self.lavoro.start()

    def _interrompi(self):
        self.stop.set()
        self.btn_stop.config(state="disabled")
        self.lbl_prog.config(text="Interruzione in corso…")

    def _svuota_coda(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "log":
                    self._scrivi_log(m[1])
                elif m[0] == "prog":
                    self._avanzamento(*m[1:])
                elif m[0] == "fine":
                    _, callback, res, err = m
                    callback(res, err)
                elif m[0] == "finito":
                    self._finito(m[1], m[2])
        except queue.Empty:
            pass
        self.after(100, self._svuota_coda)

    def _avanzamento(self, fase, n, t):
        if fase != self._fase:  # nuova fase: riparte la stima del tempo
            self._fase, self._t0, self._n0 = fase, time.monotonic(), n
        self.pbar.config(maximum=max(t, 1), value=n)
        testo = f"{fase}: {n}/{t}"
        fatti, trascorso = n - self._n0, time.monotonic() - self._t0
        if fatti >= 5 and trascorso > 3 and n < t:
            resto = trascorso / fatti * (t - n)
            testo += f" · circa {_durata(resto)} alla fine della fase"
        self.lbl_prog.config(text=testo)

    def _finito(self, out, err):
        if self.app.chiusura_richiesta:  # l'utente ha chiuso la finestra: dati già salvati
            self.app.destroy()
            return
        self.btn_avvia.config(state="normal")
        self.btn_stop.config(state="disabled")
        self._aggiorna_info_cache()
        if err:
            self.lbl_prog.config(text="Errore: nessun dato salvato")
            self._scrivi_log(f"\nERRORE: {err!r}")
            messagebox.showerror("Errore", f"Lo scaricamento non è partito:\n\n{err}")
            return
        self.app.esplora.aggiorna_elenco(select=out)
        meta = (self.app.esplora.dati or {}).get("meta", {})
        r = meta.get("riepilogo") or {}
        conteggi = (f"{r.get('corsi', 0)} corsi, {r.get('piani', 0)} piani di studio\n"
                    f"{r.get('insegnamenti', 0)} insegnamenti, {r.get('scaglioni', 0)} scaglioni, "
                    f"{r.get('lezioni_settimanali', 0)} lezioni settimanali")
        if r.get("insegnamenti_con_errore"):
            conteggi += (f"\n\n{r['insegnamenti_con_errore']} insegnamenti non sono stati letti per un errore "
                         "(li trovi nella colonna «Note»): puoi rilanciare lo scaricamento, "
                         "quelli già letti verranno presi dalla cache.")
        if r.get("insegnamenti_con_errore_sito"):
            conteggi += ("\n\nPer alcuni insegnamenti il sito oggi risponde con un suo errore. Docenti e orari "
                         "si trovano lo stesso nella scheda «3 · Cerca sul sito» → «Chi insegna».")
        if meta.get("errore"):
            self.lbl_prog.config(text="Fermato da un errore (dati parziali salvati)")
            titolo, icona = "Scaricamento incompleto", "warning"
            testo = (f"Lo scaricamento si è fermato per un errore:\n{meta['errore']}\n\n"
                     f"Dati raccolti fino a quel punto:\n{conteggi}")
        elif meta.get("interrotto") or self.stop.is_set():
            self.lbl_prog.config(text="Interrotto (dati parziali salvati)")
            titolo, icona = "Scaricamento interrotto", "info"
            testo = f"Dati raccolti fino all'interruzione:\n{conteggi}"
        else:
            self.lbl_prog.config(text="Completato")
            titolo, icona = "Scaricamento completato", "info"
            testo = conteggi
        if messagebox.askyesno(titolo, f"{testo}\n\nSalvati in:\n{out}\n\nVuoi esplorarli ed esportarli adesso?",
                               icon=icona):
            self.app.nb.select(self.app.esplora)

    def in_corso(self):
        return self.lavoro is not None and self.lavoro.is_alive()


# ============================================================ scheda 2: esplora

FILTRI = [
    ("corso", "Corso di studi"),
    ("piano_codice", "Piano"),
    ("periodo", "Periodo"),
    ("insegnamento", "Insegnamento"),
    ("scaglione", "Scaglione"),
    ("giorno", "Giorno"),
    ("aula", "Aula"),
    ("docenti", "Docente"),
]
NOMI_TABELLE = {"insegnamenti": "Insegnamenti", "scaglioni": "Scaglioni", "cognomi": "Orario per cognome",
                "lezioni": "Lezioni (orario)", "piani": "Corsi e piani"}


class SchedaEsplora(ttk.Frame):
    def __init__(self, app, parent):
        super().__init__(parent, padding=10)
        self.app = app
        self.dati = None
        self.tab = {}
        self.righe = []
        self.colonne_scelte = {k: list(v) for k, v in ex.COLONNE.items()}
        self.ordine = (None, False)
        self._costruisci()
        self.aggiorna_elenco()

    def _costruisci(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        s1 = Sezione(self, "① File di dati")
        s1.grid(row=0, column=0, sticky="ew")
        s1.columnconfigure(1, weight=1)
        lbl = ttk.Label(s1, text="File")
        lbl.grid(row=0, column=0)
        self.cb_file = ttk.Combobox(s1, state="readonly")
        self.cb_file.grid(row=0, column=1, sticky="ew", padx=6)
        self.cb_file.bind("<<ComboboxSelected>>", lambda e: self._carica_file())
        aiuto("I file di dati creati dalla scheda 1, dal più recente. Il nome contiene anno accademico, "
              "sede, data e ora dello scaricamento.", lbl, self.cb_file)
        ttk.Button(s1, text="Sfoglia…", command=self._sfoglia).grid(row=0, column=2)
        ttk.Button(s1, text="Aggiorna elenco", command=self.aggiorna_elenco).grid(row=0, column=3, padx=4)
        ttk.Button(s1, text="Apri cartella output", command=lambda: apri_percorso(OUTPUT)).grid(row=0, column=4)
        self.lbl_file = ttk.Label(s1, foreground=GRIGIO)
        self.lbl_file.grid(row=1, column=0, columnspan=5, sticky="w", pady=(4, 0))

        s2 = Sezione(self, "② Tabella e filtri")
        s2.grid(row=1, column=0, sticky="ew", pady=8)
        f = ttk.Frame(s2)
        f.grid(row=0, column=0, columnspan=8, sticky="w")
        self.var_tab = tk.StringVar(value="")
        for i, (k, t) in enumerate(NOMI_TABELLE.items()):
            rb = ttk.Radiobutton(f, text=t, value=k, variable=self.var_tab, command=self._tabella_cambiata)
            rb.grid(row=0, column=i, padx=(0, 14))
            aiuto(ex.DESCRIZIONI[k] + ".", rb)
        self.lbl_desc = ttk.Label(f, foreground=GRIGIO)
        self.lbl_desc.grid(row=1, column=0, columnspan=len(NOMI_TABELLE), sticky="w", pady=(2, 4))

        self.cb_filtri = {}
        self.lbl_filtri = {}
        for i, (col, nome) in enumerate(FILTRI):
            r, c = 1 + i // 4, (i % 4) * 2
            lbl = ttk.Label(s2, text=nome)
            lbl.grid(row=r, column=c, sticky="w", pady=3, padx=(0, 4))
            cb = ttk.Combobox(s2, state="readonly", width=34, values=[TUTTI])
            cb.set(TUTTI)
            cb.grid(row=r, column=c + 1, sticky="ew", padx=(0, 14), pady=3)
            cb.bind("<<ComboboxSelected>>", lambda e, col=col: self._filtro_cambiato(col))
            aiuto(f"Mostra solo le righe con questo valore di «{nome}». I filtri si sommano, e ogni menu propone "
                  "solo i valori ancora possibili con i filtri che lo precedono. Un menu grigio vuol dire "
                  "che la tabella scelta non ha questa colonna.", lbl, cb)
            self.cb_filtri[col], self.lbl_filtri[col] = cb, lbl
        for c in (1, 3, 5, 7):
            s2.columnconfigure(c, weight=1)
        f = ttk.Frame(s2)
        f.grid(row=3, column=0, columnspan=8, sticky="ew", pady=(4, 0))
        f.columnconfigure(1, weight=1)
        lbl = ttk.Label(f, text="Cerca testo")
        lbl.grid(row=0, column=0)
        self.var_cerca = tk.StringVar()
        e = ttk.Entry(f, textvariable=self.var_cerca)
        e.grid(row=0, column=1, sticky="ew", padx=6)
        aiuto("Tiene le righe che contengono il testo in una colonna qualsiasi, senza distinguere maiuscole. "
              "Con più parole, devono esserci tutte (es. «analisi lunedì»).", lbl, e)
        self._timer = None
        self.var_cerca.trace_add("write", lambda *a: self._rinvia_aggiornamento())
        self.var_unisci = tk.BooleanVar()
        cb = ttk.Checkbutton(f, text="Una riga sola per ciò che si ripete in più corsi (unisce i doppioni)",
                             variable=self.var_unisci, command=self._aggiorna_anteprima)
        cb.grid(row=0, column=2, padx=8)
        aiuto("Lo stesso insegnamento (o scaglione, o lezione) può comparire in più corsi e piani di studio. "
              "Con questa opzione resta una riga sola, e le colonne Corso e Piano elencano tutti i corsi "
              "separati da «|». Vale anche per i file salvati.", cb)
        b = ttk.Button(f, text="Colonne…", command=self._scegli_colonne)
        b.grid(row=0, column=3)
        aiuto("Scegli quali colonne mostrare e salvare.", b)
        b = ttk.Button(f, text="Azzera filtri", command=self._azzera_filtri)
        b.grid(row=0, column=4, padx=(6, 0))
        aiuto("Rimette tutti i filtri su «(tutti)» e svuota la ricerca.", b)
        fd = ttk.Frame(f)
        fd.grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 0))
        lbl = ttk.Label(fd, text="Solo lezioni dal")
        lbl.grid(row=0, column=0)
        self.var_dal = tk.StringVar()
        e = ttk.Entry(fd, textvariable=self.var_dal, width=12)
        e.grid(row=0, column=1, padx=6)
        self.lbl_dal = ttk.Label(fd, text="gg/mm/aaaa, vuoto = tutte", foreground=GRIGIO)
        self.lbl_dal.grid(row=0, column=2)
        aiuto("Toglie le lezioni già concluse prima di questa data, per esempio quelle delle prime settimane "
              "quando poi l'aula è cambiata. Scrivi la data come 05/10/2026; lascia vuoto per tenere tutte "
              "le lezioni. Vale per tutte le tabelle, per i file salvati e per il calendario.", lbl, e)
        self.var_dal.trace_add("write", lambda *a: self._ricalcola_tabelle())
        self.var_consecutive = tk.BooleanVar()
        cb = ttk.Checkbutton(fd, text="Unisci le lezioni consecutive", variable=self.var_consecutive,
                             command=self._ricalcola_tabelle)
        cb.grid(row=0, column=3, padx=(16, 0))
        aiuto("Due lezioni attaccate una all'altra dello stesso insegnamento, nella stessa aula e negli stessi "
              "giorni (es. 14:15–16:15 e 16:15–18:15) diventano una sola (14:15–18:15): il sito a volte registra "
              "così un'unica lezione lunga. Vale per tutte le tabelle, per i file salvati e per il calendario.", cb)

        s3 = Sezione(self, "③ Anteprima (clic su un'intestazione per ordinare)")
        s3.grid(row=3, column=0, sticky="nsew")
        s3.columnconfigure(0, weight=1)
        s3.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(s3, show="headings")
        ys = ttk.Scrollbar(s3, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(s3, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._dettaglio_riga)
        self.tree.tag_configure("alt", background="#f2f5f9")  # righe alternate, più facili da seguire
        self.lbl_n = ttk.Label(s3, text="")
        self.lbl_n.grid(row=2, column=0, sticky="w", pady=(4, 0))

        s4 = Sezione(self, "④ Salva le righe filtrate")
        s4.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        formati = [
            ("xlsx", "Excel (.xlsx)", "La tabella che vedi, in un foglio Excel con intestazioni e filtri."),
            ("csv", "CSV (.csv)", "Testo separato da «;»: si apre direttamente in Excel italiano e in "
                                  "qualunque foglio di calcolo."),
            ("html", "Pagina web (.html)", "Una pagina da aprire nel browser, con ricerca e ordinamento per colonna."),
            ("json", "JSON (.json)", "Formato per programmi: una lista di righe con le colonne scelte."),
        ]
        for i, (fmt, testo, spiegazione) in enumerate(formati):
            b = ttk.Button(s4, text=testo, command=lambda f=fmt: self._esporta(f))
            b.grid(row=0, column=i, padx=(0, 6))
            aiuto(spiegazione + " Salva solo le righe e le colonne che vedi nell'anteprima.", b)
        b = ttk.Button(s4, text="Excel con tutte le tabelle", command=self._esporta_tutte)
        b.grid(row=0, column=4, padx=(12, 6))
        aiuto("Un file Excel con un foglio per ogni tabella (Insegnamenti, Scaglioni, …), "
              "con gli stessi filtri applicati a ciascuna.", b)
        ttk.Separator(s4, orient="vertical").grid(row=0, column=5, sticky="ns", padx=10)
        lbl = ttk.Label(s4, text="Orario settimanale:")
        lbl.grid(row=0, column=6)
        self.cb_cal = ttk.Combobox(s4, state="readonly", width=46, values=list(ex.RAGGRUPPA.values()))
        self.cb_cal.set(VUOTO)
        self.cb_cal.grid(row=0, column=7, padx=6)
        b = ttk.Button(s4, text="Crea calendario", command=self._calendario)
        b.grid(row=0, column=8)
        aiuto("Una pagina web con griglie settimanali Lunedì–Venerdì, da stampare o salvare in PDF dal browser. "
              "Scegli nel menu come dividere le lezioni in griglie: per esempio «per fascia di cognomi» dà "
              "l'orario completo di uno studente. Usa le lezioni che rispettano i filtri attuali.",
              lbl, self.cb_cal, b)

    # ---------------------------------------------------------------- file
    def aggiorna_elenco(self, select=None):
        files = sorted(OUTPUT.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True) if OUTPUT.exists() else []
        self.files = {p.name: p for p in files}
        if select:
            select = Path(select)
            self.files.setdefault(select.name, select)
        self.cb_file.config(values=list(self.files))
        if select:
            self.cb_file.set(Path(select).name)
            self._carica_file()
        elif not self.cb_file.get() and not files:
            self.lbl_file.config(text="Nessun file: scarica prima i dati dalla scheda «Scarica dati dal sito».")
        elif not self.cb_file.get():
            self.lbl_file.config(text="Scegli un file di dati.")

    def _sfoglia(self):
        p = filedialog.askopenfilename(initialdir=OUTPUT if OUTPUT.exists() else HERE,
                                       filetypes=[("File dati JSON", "*.json")])
        if p:
            self.files[Path(p).name] = Path(p)
            self.cb_file.config(values=list(self.files))
            self.cb_file.set(Path(p).name)
            self._carica_file()

    def _carica_file(self):
        p = self.files.get(self.cb_file.get())
        if not p:
            return
        try:
            self.dati = ex.carica(p)
            self.tab = self._tabelle()
        except Exception as e:
            self.dati, self.tab, self.righe = None, {}, []
            self.tree.delete(*self.tree.get_children())
            self.lbl_file.config(text="Il file scelto non è leggibile.")
            messagebox.showerror("File non valido", f"{p}\n\n{e}")
            return
        n = {k: len(v) for k, v in self.tab.items()}
        corsi = len(self.dati.get("corsi_di_studio", []))
        self.lbl_file.config(text=f"{ex.descrivi_file(self.dati)}  —  {corsi} corsi, {n['insegnamenti']} insegnamenti, "
                                  f"{n['scaglioni']} scaglioni, {n['lezioni']} lezioni settimanali")
        if not self.var_tab.get():
            self.var_tab.set("scaglioni")
        self._tabella_cambiata()

    # ---------------------------------------------------------------- filtri
    def _lezioni_dal(self):
        """La data di «Solo lezioni dal», o None se vuota o non valida (lo dice l'etichetta accanto)."""
        try:
            dal = ex.leggi_data(self.var_dal.get())
        except ValueError:
            self.lbl_dal.config(text="data non valida: scrivila come 05/10/2026", foreground=ROSSO)
            return None
        self.lbl_dal.config(text=f"lezioni concluse prima del {dal:%d/%m/%Y} escluse" if dal
                            else "gg/mm/aaaa, vuoto = tutte", foreground=GRIGIO)
        return dal

    def _tabelle(self):
        return ex.tabelle(self.dati, self._lezioni_dal(), self.var_consecutive.get())

    def _ricalcola_tabelle(self):
        """«Solo lezioni dal» o «Unisci le lezioni consecutive» cambiati: le tabelle vanno ricostruite."""
        if self.dati is None:
            self._lezioni_dal()  # aggiorna comunque l'etichetta della data
            return
        self.tab = self._tabelle()
        self._aggiorna_valori_filtri()
        self._rinvia_aggiornamento()

    def _tabella_cambiata(self):
        t = self.var_tab.get()
        self.lbl_desc.config(text=ex.DESCRIZIONI.get(t, ""))
        cols = ex.COLONNE.get(t, [])
        for col, cb in self.cb_filtri.items():
            stato = "readonly" if col in cols else "disabled"
            cb.config(state=stato)
            if col not in cols:
                cb.set(TUTTI)
        self.ordine = (None, False)
        self._aggiorna_valori_filtri()
        self._aggiorna_anteprima()

    def _righe_base(self, fino_a=None):
        """Righe della tabella corrente filtrate dai filtri che precedono `fino_a`."""
        righe = self.tab.get(self.var_tab.get(), [])
        for col, _ in FILTRI:
            if col == fino_a:
                break
            v = self.cb_filtri[col].get()
            if v and v != TUTTI:
                if col == "docenti":
                    righe = [r for r in righe if v in str(r.get("docenti") or "")]
                else:
                    righe = [r for r in righe if str(r.get(col)) == v]
        return righe

    def _aggiorna_valori_filtri(self):
        cols = ex.COLONNE.get(self.var_tab.get(), [])
        for col, _ in FILTRI:
            cb = self.cb_filtri[col]
            if col not in cols:
                continue
            base = self._righe_base(fino_a=col)
            if col == "docenti":
                vals = sorted({d.strip() for r in base for d in str(r.get("docenti") or "").split(",") if d.strip()},
                              key=str.lower)
            elif col == "giorno":
                ordine = list(ex.GIORNI_BREVI)
                vals = sorted({str(r.get(col)) for r in base if r.get(col)},
                              key=lambda g: ordine.index(g) if g in ordine else 99)
            else:
                vals = ex.valori_distinti(base, col)
            cb.config(values=[TUTTI] + vals)
            if cb.get() not in vals:
                cb.set(TUTTI)

    def _filtro_cambiato(self, col):
        self._aggiorna_valori_filtri()
        self._aggiorna_anteprima()

    def _azzera_filtri(self):
        for cb in self.cb_filtri.values():
            cb.set(TUTTI)
        self.var_cerca.set("")
        self._aggiorna_valori_filtri()
        self._aggiorna_anteprima()

    def _rinvia_aggiornamento(self):
        if self._timer:
            self.after_cancel(self._timer)
        self._timer = self.after(300, self._aggiorna_anteprima)

    def _righe_filtrate(self, tabella=None, completo=False):
        """Colonne e righe da mostrare/salvare. completo=True: righe intere, senza unire i duplicati."""
        t = tabella or self.var_tab.get()
        cols = self.colonne_scelte.get(t, ex.COLONNE[t])
        if tabella and tabella != self.var_tab.get():
            # per le altre tabelle valgono solo i filtri sulle colonne che hanno
            righe = self.tab.get(t, [])
            for col, _ in FILTRI:
                v = self.cb_filtri[col].get()
                if v and v != TUTTI and col in ex.COLONNE[t]:
                    if col == "docenti":
                        righe = [r for r in righe if v in str(r.get(col) or "")]
                    else:
                        righe = [r for r in righe if str(r.get(col)) == v]
        else:
            righe = self._righe_base()
        righe = ex.filtra(righe, None, self.var_cerca.get())
        if self.var_unisci.get() and not completo:
            righe = ex.unisci_duplicati(righe, cols)
        return cols, righe

    # ---------------------------------------------------------------- anteprima
    def _aggiorna_anteprima(self):
        self._timer = None
        if not self.tab or not self.var_tab.get():
            return
        cols, righe = self._righe_filtrate()
        col_ord, desc = self.ordine
        if col_ord in cols:
            righe = sorted(righe, key=lambda r: (r.get(col_ord) is None, _chiave(r.get(col_ord))), reverse=desc)
        self.righe = righe
        self.tree.delete(*self.tree.get_children())
        self.tree.config(columns=cols)
        for c in cols:
            freccia = (" ▼" if desc else " ▲") if c == col_ord else ""
            self.tree.heading(c, text=ex.label(c) + freccia, command=lambda c=c: self._ordina(c))
            lung = max([len(ex.label(c))] + [len(str(r.get(c) or "")) for r in righe[:200]])
            self.tree.column(c, width=min(max(lung * 8 + 16, 60), 400), stretch=False)
        for n, r in enumerate(righe[:MAX_ANTEPRIMA]):
            # nell'anteprima una cella mostra una riga sola: le lezioni dello stesso giorno separate da " / "
            self.tree.insert("", "end", values=["" if r.get(c) is None else str(r.get(c)).replace(chr(10), " / ")
                                                for c in cols],
                             tags=("alt",) if n % 2 else ())
        extra = f" (nell'anteprima le prime {MAX_ANTEPRIMA}; il salvataggio le include tutte)" if len(righe) > MAX_ANTEPRIMA else ""
        self.lbl_n.config(text=f"{len(righe)} righe{extra}. Doppio clic su una riga per vederla per intero.")

    def _ordina(self, col):
        c, desc = self.ordine
        self.ordine = (col, not desc if c == col else False)
        self._aggiorna_anteprima()

    def _dettaglio_riga(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        riga = self.righe[self.tree.index(iid)]
        w = tk.Toplevel(self)
        w.title("Dettaglio riga")
        w.geometry("720x520")
        t = ScrolledText(w, wrap="word", font=("", 10))
        t.pack(fill="both", expand=True)
        for c in self.colonne_scelte[self.var_tab.get()]:
            t.insert("end", f"{ex.label(c)}: ", "b")
            t.insert("end", f"{'' if riga.get(c) is None else riga.get(c)}\n")
        t.tag_config("b", font=("", 10, "bold"))
        t.config(state="disabled")
        url = riga.get("url")
        if url and str(url).startswith("http") and " | " not in str(url):
            ttk.Button(w, text="Apri la pagina sul sito del PoliMi",
                       command=lambda: webbrowser.open(url)).pack(pady=6)

    def _scegli_colonne(self):
        t = self.var_tab.get()
        if not t:
            return
        w = tk.Toplevel(self)
        w.title("Colonne da mostrare e salvare")
        w.transient(self.winfo_toplevel())
        vars_ = {}
        fr = ttk.Frame(w, padding=10)
        fr.pack(fill="both", expand=True)
        tutte = ex.COLONNE[t]
        for i, c in enumerate(tutte):
            v = tk.BooleanVar(value=c in self.colonne_scelte[t])
            ttk.Checkbutton(fr, text=ex.label(c), variable=v).grid(row=i % 16, column=i // 16, sticky="w", padx=8)
            vars_[c] = v
        b = ttk.Frame(w, padding=10)
        b.pack(fill="x")

        def imposta(val):
            for v in vars_.values():
                v.set(val)

        def ok():
            scelte = [c for c in tutte if vars_[c].get()]
            if not scelte:
                messagebox.showwarning("Colonne", "Scegli almeno una colonna.", parent=w)
                return
            self.colonne_scelte[t] = scelte
            w.destroy()
            self._aggiorna_anteprima()
        ttk.Button(b, text="Tutte", command=lambda: imposta(True)).pack(side="left")
        ttk.Button(b, text="Nessuna", command=lambda: imposta(False)).pack(side="left", padx=6)
        ttk.Button(b, text="OK", command=ok).pack(side="right")
        w.grab_set()

    # ---------------------------------------------------------------- export
    def _pronto(self):
        if not self.tab or not self.var_tab.get():
            messagebox.showinfo("Nessun dato", "Scegli prima un file di dati e una tabella.")
            return False
        return True

    def _nome_suggerito(self, base, ext):
        parti = [base]
        for col in ("corso", "piano_codice", "insegnamento", "giorno", "aula"):
            v = self.cb_filtri[col].get()
            if v and v != TUTTI:
                parti.append(v.split("(")[0].strip()[:30])
        nome = "_".join(parti)
        nome = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in nome).strip("_")
        return f"{nome}.{ext}"

    def _chiedi_percorso(self, nome, ext, descr):
        cartella = OUTPUT / "esportazioni"
        cartella.mkdir(parents=True, exist_ok=True)
        return filedialog.asksaveasfilename(initialdir=cartella, initialfile=nome, defaultextension=f".{ext}",
                                            filetypes=[(descr, f"*.{ext}")])

    def _salva(self, path, scrivi, testo):
        salva_file(self, path, scrivi, testo)

    def _esporta(self, fmt):
        if not self._pronto():
            return
        t = self.var_tab.get()
        cols, righe = self._righe_filtrate()
        if not righe:
            messagebox.showinfo("Niente da salvare", "Con questi filtri non ci sono righe.")
            return
        p = self._chiedi_percorso(self._nome_suggerito(t, fmt), fmt, ex.FORMATI[fmt])
        if not p:
            return
        self._salva(p, lambda: ex.esporta(p, fmt, NOMI_TABELLE[t], cols, righe, ex.descrivi_file(self.dati)),
                    f"{len(righe)} righe salvate ({NOMI_TABELLE[t].lower()}).")

    def _esporta_tutte(self):
        if not self._pronto():
            return
        p = self._chiedi_percorso(self._nome_suggerito("manifesti", "xlsx"), "xlsx", "Excel")
        if not p:
            return
        fogli = {NOMI_TABELLE[t]: self._righe_filtrate(t) for t in ex.COLONNE}
        conteggi = ", ".join(f"{nome}: {len(r)}" for nome, (_, r) in fogli.items())
        self._salva(p, lambda: ex.esporta_xlsx(p, fogli),
                    f"Un foglio per ogni tabella, con gli stessi filtri applicati.\nRighe per foglio — {conteggi}.")

    def _calendario(self):
        if not self._pronto():
            return
        scelta = self.cb_cal.get()
        per = next((k for k, v in ex.RAGGRUPPA.items() if v == scelta), None)
        if not per:
            messagebox.showinfo("Orario settimanale", "Scegli prima come raggruppare le lezioni.")
            return
        # righe complete, senza «Unisci» né scelta colonne: al calendario servono giorno e ora,
        # e le lezioni ripetute in più piani le toglie da sé
        _, righe = self._righe_filtrate("lezioni", completo=True)
        righe = [r for r in righe if r.get("giorno_n") and r.get("inizio") and r.get("fine")]
        if not righe:
            messagebox.showinfo("Niente da mostrare", "Con questi filtri non ci sono lezioni con orario.\n\n"
                                "Controlla i filtri, oppure scarica i dati con l'opzione «Orario delle lezioni».")
            return
        p = self._chiedi_percorso(self._nome_suggerito(f"orario_{per}", "html"), "html", "Pagina web")
        if not p:
            return
        self._salva(p, lambda: ex.esporta_calendario(p, righe, per, sottotitolo=ex.descrivi_file(self.dati)),
                    "Orario settimanale creato: si apre nel browser e si può stampare o salvare in PDF.")


# ============================================================ scheda 3: cerca sul sito

TUTTE_SEDI = "(tutte le sedi)"
GIORNI_FASCE = ["lun", "mar", "mer", "gio", "ven", "sab"]

# le ricerche della scheda 3: chiave -> (titolo del pulsante, spiegazione con un esempio)
RICERCHE = {
    "chi": ("Chi insegna un insegnamento",
            "I docenti di un insegnamento, ognuno con il suo orario e il suo scaglione. Se scrivi delle fasce "
            "orarie, in cima trovi chi le copre tutte. Esempio: «geometria e algebra lineare», sede Milano "
            "Leonardo, fasce «gio 08:15-10:15, ven 10:15-13:15»."),
    "docente": ("Scheda di un docente",
                "Insegnamenti, scaglioni (lettere dei cognomi), orario delle lezioni e contatti di un docente. "
                "Scrivi il cognome o una parte del nome (es. «compagnoni»); se ci sono più docenti, scegli quello "
                "giusto dall'elenco con un doppio clic."),
    "insegnamenti": ("Insegnamenti e docenti",
                     "Gli insegnamenti che contengono le parole cercate (o il codice), ognuno con i suoi docenti. "
                     "Puoi cercare anche solo per docente."),
    "aule": ("Aule: chi le occupa / aule libere",
             "Per uno o più giorni: chi usa ogni aula e a che ora (insegnamento e docente), oppure quali aule sono "
             "libere in una fascia oraria. I dati sono quelli del sito «Spazi» del Politecnico."),
    "corso": ("Informazioni su un corso di studi",
              "Struttura e piani di studio, elenco dei docenti con i loro insegnamenti, programmi "
              "interdisciplinari, accordi di scambio internazionali. Serve il codice del corso (es. 531)."),
    "vecchi": ("Vecchi ordinamenti (prima del D.M. 509)",
               "Insegnamenti delle lauree di vecchio ordinamento, cercati per nome o per docente."),
}


class SchedaCerca(ttk.Frame):
    """Ricerche al volo sul sito (cerca.py): non serve scaricare nulla prima, i risultati arrivano
    in pochi secondi. Ogni ricerca gira in un thread e manda messaggi nella coda self.q."""

    def __init__(self, app, parent):
        super().__init__(parent, padding=10)
        self.app = app
        self.q = queue.Queue()
        self.stop = threading.Event()
        self.lavoro = None
        self.ris = None                 # ultimo cerca.Risultato
        self.righe = []                 # righe mostrate, nell'ordine della tabella
        self.ordine = (None, False)
        self.map_aa, self.map_sede, self.map_sede_aule = {}, {}, {}
        self.var = {}                   # campi di testo: chiave -> StringVar
        self.cb = {}                    # menu: chiave -> Combobox
        self._costruisci()
        self._tipo_cambiato()
        self._carica_scelte()
        self.after(100, self._svuota_coda)

    # ---------------------------------------------------------------- layout
    def _costruisci(self):
        # in alto, affiancati: ① il tipo di ricerca (a sinistra) e ② i suoi campi; sotto, tutta la
        # larghezza e l'altezza rimasta per ③ i risultati
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        s1 = Sezione(self, "① Cosa vuoi cercare")
        s1.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.var_tipo = tk.StringVar(value="chi")
        for i, (k, (titolo, spiegazione)) in enumerate(RICERCHE.items()):
            rb = ttk.Radiobutton(s1, text=titolo, value=k, variable=self.var_tipo, command=self._tipo_cambiato)
            rb.grid(row=i, column=0, sticky="w", pady=2)
            aiuto(spiegazione, rb)

        s2 = Sezione(self, "② Dati della ricerca")
        s2.grid(row=0, column=1, sticky="nsew")
        s2.columnconfigure(0, weight=1)
        self.lbl_spiega = ttk.Label(s2, foreground=GRIGIO, justify="left")
        self.lbl_spiega.grid(row=0, column=0, sticky="w", pady=(0, 6))
        s2.bind("<Configure>", lambda e: self.lbl_spiega.config(wraplength=max(300, e.width - 30)))
        self.fr_anno = ttk.Frame(s2)
        self.fr_anno.grid(row=1, column=0, sticky="w")
        lbl = ttk.Label(self.fr_anno, text="Anno accademico", width=24)
        lbl.grid(row=0, column=0, sticky="w")
        self.cb_aa = ttk.Combobox(self.fr_anno, state="readonly", width=14)
        self.cb_aa.grid(row=0, column=1, sticky="w", pady=3)
        aiuto("L'anno accademico in cui cercare. Parte da quello attuale del sito; l'elenco arriva dal sito.",
              lbl, self.cb_aa)

        self.moduli = {}
        for k in RICERCHE:
            fr = ttk.Frame(s2)
            fr.grid(row=2, column=0, sticky="ew")
            self.moduli[k] = fr

        # chi insegna
        fr = self.moduli["chi"]
        self._campo(fr, 0, "chi_ins", "Insegnamento", "es. geometria e algebra lineare  oppure  082747",
                    "Il nome dell'insegnamento (anche solo una parte, almeno 3 lettere) oppure il suo codice.")
        self._menu(fr, 1, "chi_sede", "Sede", "Cerca solo i docenti che insegnano in questa sede.")
        self._campo(fr, 2, "chi_fasce", "Fasce orarie (facoltative)", "es. gio 08:15-10:15, ven 10:15-13:15",
                    "Le lezioni che cerchi, separate da virgole. Un intervallo (gio 08:15-10:15) deve essere "
                    "coperto tutto da una lezione; un'ora sola (gio 08:15) vuol dire «a lezione in quel "
                    "momento». Lascia vuoto per vedere gli orari di tutti i docenti.", larghezza=50)
        aggiungi = ttk.Frame(fr)
        aggiungi.grid(row=3, column=1, columnspan=2, sticky="w", pady=(0, 4))
        self.cb_giorno = ttk.Combobox(aggiungi, state="readonly", width=5, values=GIORNI_FASCE)
        self.cb_giorno.grid(row=0, column=0)
        ttk.Label(aggiungi, text="dalle").grid(row=0, column=1, padx=4)
        self.var["f_dalle"] = tk.StringVar()
        ttk.Entry(aggiungi, textvariable=self.var["f_dalle"], width=6).grid(row=0, column=2)
        ttk.Label(aggiungi, text="alle").grid(row=0, column=3, padx=4)
        self.var["f_alle"] = tk.StringVar()
        ttk.Entry(aggiungi, textvariable=self.var["f_alle"], width=6).grid(row=0, column=4)
        b = ttk.Button(aggiungi, text="＋ Aggiungi fascia", command=self._aggiungi_fascia)
        b.grid(row=0, column=5, padx=(8, 4))
        aiuto("Un modo guidato per scrivere le fasce: scegli il giorno, scrivi le ore (es. 08:15 e 10:15; "
              "«alle» si può lasciare vuoto) e premi qui. La fascia si aggiunge al campo sopra.", b, self.cb_giorno)
        ttk.Button(aggiungi, text="Svuota", command=lambda: self.var["chi_fasce"].set("")).grid(row=0, column=6)

        # scheda docente
        fr = self.moduli["docente"]
        self._campo(fr, 0, "doc_chi", "Docente", "es. compagnoni  oppure il codice  245289",
                    "Il cognome o una parte del nome (almeno 3 lettere), oppure il codice numerico del docente.")

        # insegnamenti e docenti
        fr = self.moduli["insegnamenti"]
        self._campo(fr, 0, "ins_testo", "Insegnamento", "es. analisi matematica 1  oppure  082740",
                    "Parte del nome o del codice dell'insegnamento. Puoi lasciarlo vuoto se cerchi per docente.")
        self._campo(fr, 1, "ins_doc", "Docente (facoltativo)", "es. rossi",
                    "Parte del nome del docente. Puoi lasciarlo vuoto se cerchi per insegnamento.")
        self._menu(fr, 2, "ins_sede", "Sede", "Cerca solo gli insegnamenti di questa sede.")

        # aule
        fr = self.moduli["aule"]
        self._menu(fr, 0, "aule_sede", "Sede", "La sede (o il singolo indirizzo) di cui vedere le aule. "
                   "Milano Leonardo è «Milano Città Studi».")
        e = self._campo(fr, 1, "aule_giorno", "Giorno", "",
                        "Il giorno da guardare, scritto come 15/10/2026.", larghezza=12)
        rapidi = ttk.Frame(fr)
        rapidi.grid(row=1, column=2, sticky="w", padx=8)
        for testo, giorni in (("Oggi", 0), ("Domani", 1)):
            ttk.Button(rapidi, text=testo, width=8, command=lambda g=giorni: self.var["aule_giorno"].set(
                (date.today() + timedelta(days=g)).strftime("%d/%m/%Y"))).pack(side="left", padx=(0, 4))
        ttk.Label(rapidi, text="gg/mm/aaaa", foreground=GRIGIO).pack(side="left", padx=4)
        self._campo(fr, 2, "aule_al", "Fino al (facoltativo)", "per più giorni insieme, al massimo 14",
                    "Per guardare più giorni di fila: l'ultimo giorno, come 16/10/2026.", larghezza=12)
        self._campo(fr, 3, "aule_aula", "Aula (facoltativa)", "es. T.2.2",
                    "Solo questa aula, scritta come sul sito (es. T.2.2, 5.0.1, B.4.4).", larghezza=12)
        self._campo(fr, 4, "aule_testo", "Parole (facoltative)", "es. geometria  oppure  compagnoni  oppure  082747",
                    "Tiene solo le occupazioni che contengono tutte queste parole: nome o codice "
                    "dell'insegnamento, nome del docente.")
        ore = ttk.Frame(fr)
        ore.grid(row=5, column=1, columnspan=2, sticky="w", pady=3)
        lbl = ttk.Label(fr, text="Ore (facoltative)")
        lbl.grid(row=5, column=0, sticky="w")
        ttk.Label(ore, text="dalle").pack(side="left")
        self.var["aule_dalle"] = tk.StringVar()
        e1 = ttk.Entry(ore, textvariable=self.var["aule_dalle"], width=6)
        e1.pack(side="left", padx=4)
        ttk.Label(ore, text="alle").pack(side="left")
        self.var["aule_alle"] = tk.StringVar()
        e2 = ttk.Entry(ore, textvariable=self.var["aule_alle"], width=6)
        e2.pack(side="left", padx=4)
        ttk.Label(ore, text="es. dalle 10:15 alle 12:15", foreground=GRIGIO).pack(side="left", padx=8)
        aiuto("Una fascia oraria: con «Chi occupa» mostra le occupazioni in quelle ore; con «Aule libere» le aule "
              "libere per tutta la fascia (senza ore: gli intervalli liberi tra le 8 e le 20).", lbl, e1, e2)
        self.var_libere = tk.BooleanVar(value=False)
        modo = ttk.Frame(fr)
        modo.grid(row=6, column=1, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Radiobutton(modo, text="Chi occupa le aule", value=False, variable=self.var_libere).pack(side="left")
        ttk.Radiobutton(modo, text="Aule libere", value=True, variable=self.var_libere).pack(side="left", padx=12)
        for e in (e1, e2):
            e.bind("<Return>", lambda ev: self._cerca())

        # corso
        fr = self.moduli["corso"]
        self._campo(fr, 0, "corso_cod", "Codice del corso", "es. 531 (Ingegneria Informatica)",
                    "Il codice numerico del corso di studi. Lo trovi nella scheda 1, nella colonna «Codice» "
                    "accanto al nome del corso.", larghezza=10)
        lbl = ttk.Label(fr, text="Cosa mostrare")
        lbl.grid(row=1, column=0, sticky="w")
        pagine = ttk.Frame(fr)
        pagine.grid(row=1, column=1, columnspan=2, sticky="w", pady=3)
        self.var_pagina = tk.StringVar(value="struttura")
        for k, (nome, _) in cerca.PAGINE_CORSO.items():
            ttk.Radiobutton(pagine, text=nome, value=k, variable=self.var_pagina).pack(side="left", padx=(0, 12))
        aiuto("La pagina del sito da leggere per questo corso.", lbl)

        # vecchi ordinamenti
        fr = self.moduli["vecchi"]
        self._campo(fr, 0, "vo_ins", "Insegnamento", "es. geometria", "Parte del nome o del codice.")
        self._campo(fr, 1, "vo_doc", "Docente", "es. lella", "Parte del nome del docente.")

        piede = ttk.Frame(s2)
        piede.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        piede.columnconfigure(3, weight=1)
        self.btn_cerca = ttk.Button(piede, text="🔍  Cerca", style="Accent.TButton", command=self._cerca)
        self.btn_cerca.grid(row=0, column=0)
        aiuto("Avvia la ricerca sul sito (anche con il tasto Invio in un campo). Di solito bastano pochi "
              "secondi; «Chi insegna» legge la scheda di ogni docente e può volerci mezzo minuto.", self.btn_cerca)
        self.btn_stop = ttk.Button(piede, text="Interrompi", command=self._interrompi, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=6)
        self.barra = ttk.Progressbar(piede, mode="indeterminate", length=140)
        self.barra.grid(row=0, column=2, padx=(4, 10))
        self.barra.grid_remove()  # visibile solo mentre una ricerca è in corso
        self.lbl_stato = ttk.Label(piede, text="Collegamento al sito…", foreground=GRIGIO)
        self.lbl_stato.grid(row=0, column=3, sticky="w")

        s3 = Sezione(self, "③ Risultati (clic su un'intestazione per ordinare, doppio clic su una riga per i dettagli)")
        s3.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        s3.columnconfigure(0, weight=1)
        s3.rowconfigure(3, weight=1)
        self.lbl_titolo = ttk.Label(s3, text="Nessuna ricerca ancora.", font=("", 10, "bold"))
        self.lbl_titolo.grid(row=0, column=0, columnspan=2, sticky="w")
        self.lbl_note = ttk.Label(s3, justify="left")
        self.lbl_note.grid(row=1, column=0, columnspan=2, sticky="w")
        s3.bind("<Configure>", lambda e: self.lbl_note.config(wraplength=max(300, e.width - 30)))
        self.fr_tab = ttk.Frame(s3)
        self.fr_tab.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 2))
        self.var_tab = tk.IntVar(value=0)
        self.tree = ttk.Treeview(s3, show="headings")
        ys = ttk.Scrollbar(s3, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(s3, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=3, column=0, sticky="nsew")
        ys.grid(row=3, column=1, sticky="ns")
        xs.grid(row=4, column=0, sticky="ew")
        self.tree.tag_configure("alt", background="#f2f5f9")
        self.tree.tag_configure("pieno", background="#dff3e4")  # chi copre tutte le fasce
        self.tree.bind("<Double-1>", self._dettaglio_riga)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._aggiorna_pulsanti())
        basso = ttk.Frame(s3)
        basso.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        basso.columnconfigure(0, weight=1)
        self.lbl_n = ttk.Label(basso, text="")
        self.lbl_n.grid(row=0, column=0, sticky="w")
        self.btn_docente = ttk.Button(basso, text="Apri la scheda del docente", state="disabled",
                                      command=lambda: self._apri_docente(self._riga_selezionata()))
        self.btn_docente.grid(row=0, column=1, padx=6)
        aiuto("Cerca la scheda completa (insegnamenti, scaglioni, orario) del docente della riga selezionata.",
              self.btn_docente)
        self.btn_sito = ttk.Button(basso, text="Apri sul sito del PoliMi", state="disabled",
                                   command=lambda: self._apri_sito(self._riga_selezionata()))
        self.btn_sito.grid(row=0, column=2)

        s4 = Sezione(self, "④ Salva i risultati")
        s4.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        formati = [
            ("xlsx", "Excel (.xlsx)", "Tutte le tabelle del risultato, un foglio ciascuna."),
            ("csv", "CSV (.csv)", "La tabella che vedi, in testo separato da «;» (si apre in Excel)."),
            ("html", "Pagina web (.html)", "La tabella che vedi, in una pagina con ricerca e ordinamento."),
            ("json", "JSON (.json)", "La tabella che vedi, per altri programmi."),
        ]
        for i, (fmt, testo, spiegazione) in enumerate(formati):
            b = ttk.Button(s4, text=testo, command=lambda f=fmt: self._salva(f))
            b.grid(row=0, column=i, padx=(0, 6))
            aiuto(spiegazione, b)

    def _campo(self, fr, r, chiave, etichetta, esempio, spiegazione, larghezza=40):
        lbl = ttk.Label(fr, text=etichetta, width=24)
        lbl.grid(row=r, column=0, sticky="w")
        self.var[chiave] = tk.StringVar()
        e = ttk.Entry(fr, textvariable=self.var[chiave], width=larghezza)
        e.grid(row=r, column=1, sticky="w", pady=3)
        e.bind("<Return>", lambda ev: self._cerca())
        if esempio:
            ttk.Label(fr, text=esempio, foreground=GRIGIO).grid(row=r, column=2, sticky="w", padx=8)
        aiuto(spiegazione, lbl, e)
        return e

    def _menu(self, fr, r, chiave, etichetta, spiegazione):
        lbl = ttk.Label(fr, text=etichetta, width=24)
        lbl.grid(row=r, column=0, sticky="w")
        cb = self.cb[chiave] = ttk.Combobox(fr, state="readonly", width=38, values=[VUOTO])
        cb.set(VUOTO)
        cb.grid(row=r, column=1, sticky="w", pady=3)
        aiuto(spiegazione, lbl, cb)
        return cb

    def _tipo_cambiato(self):
        k = self.var_tipo.get()
        for nome, fr in self.moduli.items():
            if nome == k:
                fr.grid()
            else:
                fr.grid_remove()
        if k in ("aule", "vecchi"):  # il sito Spazi e i vecchi ordinamenti non dipendono dall'anno
            self.fr_anno.grid_remove()
        else:
            self.fr_anno.grid()
        self.lbl_spiega.config(text=RICERCHE[k][1])

    def _aggiungi_fascia(self):
        g, dalle, alle = self.cb_giorno.get(), self.var["f_dalle"].get().strip(), self.var["f_alle"].get().strip()
        try:
            if not g:
                raise ValueError("Scegli il giorno.")
            if not dalle:
                raise ValueError("Scrivi l'ora di inizio, per esempio 08:15.")
            fascia = f"{g} {dalle}" + (f"-{alle}" if alle else "")
            cerca.leggi_fasce(fascia)  # controlla ore e ordine
        except ValueError as e:
            messagebox.showwarning("Fascia oraria", str(e))
            return
        attuale = self.var["chi_fasce"].get().strip().rstrip(",")
        self.var["chi_fasce"].set(f"{attuale}, {fascia}" if attuale else fascia)
        self.var["f_dalle"].set(self.var["f_alle"].get())  # la fascia dopo parte spesso dove finisce questa
        self.var["f_alle"].set("")

    # ---------------------------------------------------------------- scelte dal sito
    def _carica_scelte(self):
        def run():
            try:
                self.q.put(("scelte", cerca.scelte(), None))
            except Exception as e:
                self.q.put(("scelte", None, e))
        threading.Thread(target=run, daemon=True).start()

    def _scelte_caricate(self, sc, err):
        if err:
            self.lbl_stato.config(text="Sito non raggiungibile: controlla la connessione e riavvia il programma.",
                                  foreground=ROSSO)
            return
        self.map_aa = {nome: cod for cod, nome in sc["aa"]}
        self.cb_aa.config(values=list(self.map_aa))
        attuale = next((n for n, c in self.map_aa.items() if c == sc["aa_attuale"]), None)
        if attuale:
            self.cb_aa.set(attuale)
        self.map_sede = {nome: cod for cod, nome in sc["sedi"]}
        for k in ("chi_sede", "ins_sede"):
            self.cb[k].config(values=[TUTTE_SEDI] + list(self.map_sede))
            self.cb[k].set(TUTTE_SEDI)
        self.map_sede_aule = {nome: cod for cod, nome in sc["sedi_aule"]}
        self.cb["aule_sede"].config(values=list(self.map_sede_aule))
        self.lbl_stato.config(text="Pronto: scegli cosa cercare, compila i campi e premi «Cerca».", foreground=GRIGIO)

    # ---------------------------------------------------------------- ricerca
    def _prepara(self):
        """Controlla i campi e restituisce (funzione da eseguire nel thread, descrizione per l'utente).
        Se manca qualcosa solleva ValueError con il messaggio da mostrare.
        Tutti i valori vengono letti qui: il thread della ricerca non deve toccare i widget."""
        k = self.var_tipo.get()
        v = {c: s.get().strip() for c, s in self.var.items()}
        aa = self.map_aa.get(self.cb_aa.get())
        if k not in ("aule", "vecchi") and not aa:
            raise ValueError("Aspetta che il programma si colleghi al sito: l'anno accademico non è ancora caricato.")
        sedi = {c: self.map_sede.get(self.cb[c].get()) for c in ("chi_sede", "ins_sede")}  # None = tutte
        prog = lambda n, t: self.q.put(("prog", f"Leggo le schede dei docenti: {n} di {t}…"))  # noqa: E731
        stop = self.stop
        if k == "chi":
            if len(v["chi_ins"]) < 3:
                raise ValueError("Scrivi il nome o il codice dell'insegnamento (almeno 3 lettere).")
            cerca.leggi_fasce(v["chi_fasce"])
            return (lambda: cerca.chi_insegna(v["chi_ins"], v["chi_fasce"], aa, sedi["chi_sede"], stop=stop,
                                              avanzamento=prog), "Cerco i docenti dell'insegnamento…")
        if k == "docente":
            if len(v["doc_chi"]) < 3 and not v["doc_chi"].isdigit():
                raise ValueError("Scrivi il cognome del docente (almeno 3 lettere) oppure il suo codice.")
            return lambda: cerca.scheda_docente(v["doc_chi"], aa, stop=stop), "Cerco il docente…"
        if k == "insegnamenti":
            if len(v["ins_testo"]) < 3 and len(v["ins_doc"]) < 3:
                raise ValueError("Scrivi almeno 3 lettere dell'insegnamento o del docente.")
            return (lambda: cerca.insegnamenti(v["ins_testo"], v["ins_doc"], aa, sedi["ins_sede"], stop=stop),
                    "Cerco gli insegnamenti…")
        if k == "aule":
            sede_aule = self.map_sede_aule.get(self.cb["aule_sede"].get())
            libere = self.var_libere.get()
            if not sede_aule:
                raise ValueError("Scegli la sede delle aule.")
            if not v["aule_giorno"]:
                raise ValueError("Scrivi il giorno (gg/mm/aaaa) o premi «Oggi».")
            for c in ("aule_dalle", "aule_alle"):
                if v[c]:
                    cerca._minuti(v[c])
            prog_g = lambda n, t: self.q.put(("prog", f"Leggo le occupazioni: giorno {n} di {t}…"))  # noqa: E731
            return (lambda: cerca.occupazione_aule(
                v["aule_giorno"], v["aule_al"] or None, sede_aule, v["aule_aula"] or None, v["aule_testo"] or None,
                v["aule_dalle"] or None, v["aule_alle"] or None, libere, stop=stop,
                avanzamento=prog_g), "Leggo l'occupazione delle aule…")
        if k == "corso":
            if not v["corso_cod"].isdigit():
                raise ValueError("Scrivi il codice numerico del corso (es. 531). Lo trovi nella scheda 1.")
            pagina = self.var_pagina.get()
            return (lambda: cerca.info_corso(v["corso_cod"], pagina, aa, stop=stop),
                    "Leggo la pagina del corso…")
        if len(v["vo_ins"]) < 3 and len(v["vo_doc"]) < 3:
            raise ValueError("Scrivi almeno 3 lettere dell'insegnamento o del docente.")
        return lambda: cerca.vecchi_ordinamenti(v["vo_ins"], v["vo_doc"], stop=stop), "Cerco nei vecchi ordinamenti…"

    def _cerca(self):
        if self.in_corso():
            return
        try:
            funzione, descrizione = self._prepara()
        except ValueError as e:
            messagebox.showwarning("Manca qualcosa", str(e))
            return
        self.stop.clear()
        self.btn_cerca.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.barra.grid()
        self.barra.start(12)
        self.lbl_stato.config(text=descrizione, foreground=GRIGIO)

        def run():
            try:
                self.q.put(("fine", funzione(), None))
            except Exception as e:
                self.q.put(("fine", None, e))
        self.lavoro = threading.Thread(target=run, daemon=True)
        self.lavoro.start()

    def _interrompi(self):
        self.stop.set()
        self.lbl_stato.config(text="Interrompo…")

    def in_corso(self):
        return self.lavoro is not None and self.lavoro.is_alive()

    def _svuota_coda(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "scelte":
                    self._scelte_caricate(msg[1], msg[2])
                elif msg[0] == "prog":
                    self.lbl_stato.config(text=msg[1])
                elif msg[0] == "fine":
                    self._finita(msg[1], msg[2])
        except queue.Empty:
            pass
        self.after(100, self._svuota_coda)

    def _finita(self, ris, err):
        self.btn_cerca.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.barra.stop()
        self.barra.grid_remove()
        if isinstance(err, ps.Interrotto) or (err is None and self.stop.is_set() and ris is None):
            self.lbl_stato.config(text="Ricerca interrotta.", foreground=GRIGIO)
            return
        if err:
            if isinstance(err, ValueError):
                testo = str(err)
            elif isinstance(err, ps.ErroreSito):
                testo = f"Il sito del Politecnico non risponde a questa ricerca: {err}.\n\nÈ un problema del sito: riprova più tardi."
            else:
                testo = f"La ricerca non è riuscita. Controlla la connessione a internet.\n\n{err}"
            self.lbl_stato.config(text="Ricerca non riuscita.", foreground=ROSSO)
            messagebox.showerror("Ricerca non riuscita", testo)
            return
        self.lbl_stato.config(text=f"Fatto: {ris.n_righe} righe trovate.", foreground=VERDE)
        self._mostra(ris)

    # ---------------------------------------------------------------- risultati
    def _mostra(self, ris):
        self.ris = ris
        self.lbl_titolo.config(text=ris.titolo)
        self.lbl_note.config(text="\n".join("• " + n for n in ris.note),
                             foreground=VERDE if ris.note and ris.note[0].startswith("Copre tutte") else "")
        for w in self.fr_tab.winfo_children():
            w.destroy()
        # si apre la prima tabella con dei risultati (es. la scheda di un docente senza orario pubblicato)
        self.var_tab.set(next((i for i, t in enumerate(ris.tabelle) if t.righe), 0))
        if len(ris.tabelle) > 1:
            ttk.Label(self.fr_tab, text="Tabella:").pack(side="left")
            for i, t in enumerate(ris.tabelle):
                ttk.Radiobutton(self.fr_tab, text=f"{t.nome} ({len(t.righe)})", value=i, variable=self.var_tab,
                                command=self._aggiorna_tabella).pack(side="left", padx=(8, 4))
        self.ordine = (None, False)
        self._aggiorna_tabella()

    def _tabella(self):
        if not self.ris or not self.ris.tabelle:
            return None
        return self.ris.tabelle[min(self.var_tab.get(), len(self.ris.tabelle) - 1)]

    def _aggiorna_tabella(self):
        t = self._tabella()
        self.tree.delete(*self.tree.get_children())
        if t is None:
            self.tree.config(columns=())
            self.lbl_n.config(text="")
            self._aggiorna_pulsanti()
            return
        righe = list(t.righe)
        col_ord, desc = self.ordine
        if col_ord in t.colonne:
            righe.sort(key=lambda r: (r.get(col_ord) is None, _chiave(r.get(col_ord))), reverse=desc)
        self.righe = righe
        visibili = [c for c in t.colonne if not c.startswith("url")]  # i link si aprono col pulsante
        self.tree.config(columns=t.colonne, displaycolumns=visibili)
        for c in t.colonne:
            freccia = (" ▼" if desc else " ▲") if c == col_ord else ""
            self.tree.heading(c, text=ex.label(c) + freccia, command=lambda c=c: self._ordina(c))
            lung = max([len(ex.label(c))] + [len(str(r.get(c) or "")) for r in righe[:200]])
            self.tree.column(c, width=min(max(lung * 8 + 16, 60), 480), stretch=False)
        for n, r in enumerate(righe):
            tag = "alt" if n % 2 else ""
            corr = str(r.get("corrispondenze") or "")
            if corr and "/" in corr and corr.split("/")[0] == corr.split("/")[1] != "0":
                tag = "pieno"
            self.tree.insert("", "end", values=["" if r.get(c) is None else str(r.get(c)).replace("\n", " / ")
                                                for c in t.colonne], tags=(tag,) if tag else ())
        self.lbl_n.config(text=f"{len(righe)} righe." + (" Doppio clic su una riga per vederla per intero."
                                                         if righe else ""))
        self._aggiorna_pulsanti()

    def _ordina(self, col):
        c, desc = self.ordine
        self.ordine = (col, not desc if c == col else False)
        self._aggiorna_tabella()

    def _riga_selezionata(self):
        sel = self.tree.selection()
        return self.righe[self.tree.index(sel[0])] if sel else None

    def _url(self, riga):
        if not riga:
            return None
        for c in ("url_docente", "valore"):
            u = str(riga.get(c) or "")
            if u.startswith("http"):
                return u
        return None

    def _aggiorna_pulsanti(self):
        r = self._riga_selezionata()
        self.btn_docente.config(state="normal" if r and (r.get("codice_docente") or r.get("docente")) else "disabled")
        self.btn_sito.config(state="normal" if self._url(r) else "disabled")

    def _apri_docente(self, riga):
        """La scheda del docente della riga: per codice, oppure per nome (l'occupazione delle aule ha solo il nome)."""
        chi = (riga or {}).get("codice_docente") or (riga or {}).get("docente")
        if not chi:
            return
        self.var_tipo.set("docente")
        self._tipo_cambiato()
        self.var["doc_chi"].set(str(chi))
        self._cerca()

    def _apri_sito(self, riga):
        u = self._url(riga)
        if u:
            webbrowser.open(u)

    def _dettaglio_riga(self, event):
        iid = self.tree.identify_row(event.y)
        t = self._tabella()
        if not iid or t is None:
            return
        riga = self.righe[self.tree.index(iid)]
        w = tk.Toplevel(self)
        w.title("Dettaglio riga")
        w.geometry("720x480")
        testo = ScrolledText(w, wrap="word", font=("", 10))
        testo.pack(fill="both", expand=True)
        testo.tag_config("b", font=("", 10, "bold"))
        for c in t.colonne:
            if riga.get(c) not in (None, ""):
                testo.insert("end", f"{ex.label(c)}: ", "b")
                testo.insert("end", f"{riga.get(c)}\n")
        testo.config(state="disabled")
        pulsanti = ttk.Frame(w, padding=6)
        pulsanti.pack()
        if riga.get("codice_docente") or riga.get("docente"):
            ttk.Button(pulsanti, text="Apri la scheda di questo docente",
                       command=lambda: (w.destroy(), self._apri_docente(riga))).pack(side="left", padx=4)
        if self._url(riga):
            ttk.Button(pulsanti, text="Apri sul sito del PoliMi",
                       command=lambda: self._apri_sito(riga)).pack(side="left", padx=4)
        ttk.Button(pulsanti, text="Chiudi", command=w.destroy).pack(side="left", padx=4)

    # ---------------------------------------------------------------- salva
    def _salva(self, fmt):
        t = self._tabella()
        if t is None or not any(x.righe for x in self.ris.tabelle):
            messagebox.showinfo("Niente da salvare", "Fai prima una ricerca che trovi qualcosa.")
            return
        if fmt != "xlsx" and not t.righe:
            messagebox.showinfo("Niente da salvare", "La tabella che vedi è vuota: scegline un'altra.")
            return
        cartella = OUTPUT / "ricerche"
        cartella.mkdir(parents=True, exist_ok=True)
        nome = "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in self.ris.titolo)[:70].strip()
        path = filedialog.asksaveasfilename(initialdir=cartella, initialfile=f"{nome}.{fmt}",
                                            defaultextension=f".{fmt}", filetypes=[(ex.FORMATI[fmt], f"*.{fmt}")])
        if not path:
            return
        if fmt == "xlsx":
            salva_file(self, path, lambda: cerca.salva(self.ris, path, "xlsx"),
                       "Un foglio per ogni tabella: " + ", ".join(f"{x.nome} ({len(x.righe)})"
                                                                  for x in self.ris.tabelle if x.righe) + ".")
        else:
            salva_file(self, path, lambda: ex.esporta(path, fmt, t.nome, t.colonne, t.righe, self.ris.titolo),
                       f"{len(t.righe)} righe salvate ({t.nome}).")


def _durata(secondi):
    minuti = round(secondi / 60)
    if minuti < 1:
        return "meno di un minuto"
    if minuti < 60:
        return f"{minuti} min"
    return f"{minuti // 60} h {minuti % 60:02d} min"


def _chiave(v):
    if isinstance(v, (int, float)):
        return (0, v, "")
    s = str(v)
    try:
        return (0, float(s.replace(",", ".")), "")
    except ValueError:
        return (1, 0, s.lower())


# ============================================================ finestra

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PoliMi – manifesti, orari e docenti")
        self.chiusura_richiesta = False
        self.protocol("WM_DELETE_WINDOW", self._chiudi)
        self._stile()
        w, h = min(1400, self.winfo_screenwidth() - 60), min(900, self.winfo_screenheight() - 100)
        self.geometry(f"{w}x{h}+20+20")
        self.minsize(1000, 640)

        # in alto: cosa fa il programma, come si usa, dove trovare la guida
        testa = ttk.Frame(self, padding=(12, 8, 12, 0))
        testa.pack(fill="x")
        testa.columnconfigure(0, weight=1)
        ttk.Label(testa, text=COSA_FA_BREVE, wraplength=w - 200, justify="left",
                  font=("", 10, "bold")).grid(row=0, column=0, sticky="w")
        passi = ("Come si usa: 1 · scegli cosa scaricare e avvia  →  2 · consulta i dati e salvali.   "
                 "Per una domanda precisa (es. chi insegna Geometria il giovedì alle 8:15): 3 · Cerca sul sito.   "
                 "Mouse fermo su un'opzione = spiegazione.")
        if self._ci_sono_dati():
            passi += "   Hai già dei dati: puoi andare subito alla scheda 2."
        ttk.Label(testa, text=passi, foreground=GRIGIO, wraplength=w - 200,
                  justify="left").grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Button(testa, text="❓ Guida", command=lambda: mostra_guida(self)).grid(row=0, column=1, rowspan=2,
                                                                                  padx=(12, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.scarica = SchedaScarica(self, self.nb)
        self.esplora = SchedaEsplora(self, self.nb)
        self.cerca = SchedaCerca(self, self.nb)
        self.nb.add(self.scarica, text="  1 · Scarica dati dal sito  ")
        self.nb.add(self.esplora, text="  2 · Esplora ed esporta  ")
        self.nb.add(self.cerca, text="  3 · Cerca sul sito  ")

    @staticmethod
    def _ci_sono_dati():
        return OUTPUT.exists() and any(OUTPUT.glob("*.json"))

    def _stile(self):
        st = ttk.Style(self)
        if sys.platform.startswith("win"):
            st.theme_use("vista" if "vista" in st.theme_names() else st.theme_use())
        elif "clam" in st.theme_names():
            st.theme_use("clam")
        st.configure("Accent.TButton", font=("", 10, "bold"))
        # altezza righe dal font, così il testo non viene tagliato con lo zoom dello schermo al 125-150%
        st.configure("Treeview", rowheight=tkfont.nametofont("TkDefaultFont").metrics("linespace") + 8)

    def _chiudi(self):
        if not self.scarica.in_corso():
            self.destroy()
            return
        if messagebox.askyesno("Scaricamento in corso",
                               "Uno scaricamento è ancora in corso.\n\n"
                               "Vuoi interromperlo e chiudere? I dati raccolti finora verranno salvati."):
            self.chiusura_richiesta = True
            self.scarica._interrompi()  # alla fine del salvataggio _finito chiude la finestra


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
