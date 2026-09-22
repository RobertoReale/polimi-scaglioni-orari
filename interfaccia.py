#!/usr/bin/env python3
"""
Interfaccia grafica dello scraper dei Manifesti degli Studi PoliMi.
Avvio: avvia.bat (Windows), ./avvia.sh (Linux/macOS) oppure  python avvia.py
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
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from tkinter.scrolledtext import ScrolledText

import esporta as ex
import scarica_manifesti as ps

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "output"
CACHE = HERE / "cache"
VUOTO = "— scegli —"
TUTTI_TIPI = "Tutti i tipi"


def tipo_laurea(gruppo):
    """'Laurea Magistrale - ord. 96/23' -> 'Laurea Magistrale' (il tipo senza l'ordinamento)."""
    return re.split(r"\s+-\s+ord\b", gruppo or "", maxsplit=1, flags=re.I)[0].strip() or "Altro"
TUTTI = "(tutti)"
ON, OFF, MEZZO = "☑", "☐", "◩"
MAX_ANTEPRIMA = 3000


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
    def __init__(self, parent, titolo, **kw):
        super().__init__(parent, text=f" {titolo} ", padding=(10, 6), **kw)


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
        self.rowconfigure(1, weight=1)

        # ① anno e sede
        s1 = Sezione(self, "① Anno accademico e sede")
        s1.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(s1, text="Anno accademico").grid(row=0, column=0, sticky="w")
        self.cb_aa = ttk.Combobox(s1, state="readonly", width=14, values=[VUOTO])
        self.cb_aa.grid(row=0, column=1, padx=(6, 18))
        ttk.Label(s1, text="Sede").grid(row=0, column=2, sticky="w")
        self.cb_sede = ttk.Combobox(s1, state="readonly", width=26, values=[VUOTO])
        self.cb_sede.grid(row=0, column=3, padx=(6, 18))
        ttk.Label(s1, text="Tipo di laurea").grid(row=0, column=4, sticky="w")
        self.cb_tipo = ttk.Combobox(s1, state="readonly", width=30, values=[TUTTI_TIPI])
        self.cb_tipo.set(TUTTI_TIPI)
        self.cb_tipo.grid(row=0, column=5, padx=(6, 18))
        self.cb_tipo.bind("<<ComboboxSelected>>", lambda e: self._riempi_albero())
        self.lbl_stato = ttk.Label(s1, text="Collegamento al sito…", foreground="#666")
        self.lbl_stato.grid(row=0, column=6, sticky="w")
        for cb in (self.cb_aa, self.cb_sede):
            cb.bind("<<ComboboxSelected>>", lambda e: self._carica_corsi())

        # ② corsi
        s2 = Sezione(self, "② Corsi di studio (clic per selezionare; clic su un gruppo = tutto il gruppo)")
        s2.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        s2.columnconfigure(1, weight=1)
        s2.rowconfigure(1, weight=1)
        ttk.Label(s2, text="Cerca").grid(row=0, column=0, sticky="w")
        self.var_cerca = tk.StringVar()
        self.var_cerca.trace_add("write", lambda *a: self._riempi_albero())
        ttk.Entry(s2, textvariable=self.var_cerca).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(s2, text="Seleziona visibili", command=lambda: self._seleziona_visibili(True)).grid(row=0, column=2)
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
        s3 = Sezione(self, "③ Cosa includere")
        s3.grid(row=1, column=1, sticky="nsew")
        s3.columnconfigure(0, weight=1)
        r = 0

        ttk.Label(s3, text="Anni di corso", font=("", 9, "bold")).grid(row=r, column=0, sticky="w"); r += 1
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        self.var_anni = {}
        for i, a in enumerate(["1", "2", "3", "4", "5", "6"]):
            v = tk.BooleanVar()
            ttk.Checkbutton(f, text=f"{a}°", variable=v, command=self._anni_cambiati).grid(row=0, column=i, padx=(0, 6))
            self.var_anni[a] = v
        self.var_anni_tutti = tk.BooleanVar()
        ttk.Checkbutton(f, text="Tutti insieme", variable=self.var_anni_tutti,
                        command=self._anni_tutti_cambiato).grid(row=0, column=6, padx=(6, 0))

        ttk.Label(s3, text="Periodo didattico", font=("", 9, "bold")).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        self.var_periodi = {}
        for i, (k, t) in enumerate(ps.PERIODI.items()):
            v = tk.BooleanVar()
            ttk.Checkbutton(f, text=t, variable=v).grid(row=0, column=i, padx=(0, 8))
            self.var_periodi[k] = v

        ttk.Label(s3, text="Piani di studio", font=("", 9, "bold")).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        self.var_piani = tk.StringVar(value="")
        ttk.Radiobutton(s3, text="Tutti i piani della sede", value="tutti", variable=self.var_piani).grid(row=r, column=0, sticky="w"); r += 1
        ttk.Radiobutton(s3, text="Solo il primo piano di ogni corso", value="primo", variable=self.var_piani).grid(row=r, column=0, sticky="w"); r += 1
        self.var_altre_sedi = tk.BooleanVar()
        ttk.Checkbutton(s3, text="Tieni anche i piani erogati in altre sedi (es. Cremona)",
                        variable=self.var_altre_sedi).grid(row=r, column=0, sticky="w"); r += 1
        self.var_nondiv = tk.BooleanVar()
        ttk.Checkbutton(s3, text="Includi l'offerta non diversificata (piano ***)",
                        variable=self.var_nondiv).grid(row=r, column=0, sticky="w"); r += 1

        ttk.Label(s3, text="Dettagli da scaricare", font=("", 9, "bold")).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        self.var_scaglioni = tk.BooleanVar()
        self.var_orari = tk.BooleanVar()
        ttk.Checkbutton(s3, text="Scaglioni (lettere, docenti, moduli)", variable=self.var_scaglioni,
                        command=self._dettagli_cambiati).grid(row=r, column=0, sticky="w"); r += 1
        ttk.Checkbutton(s3, text="Orario delle lezioni (giorni, ore, aule)", variable=self.var_orari,
                        command=lambda: self.var_orari.get() and self.var_scaglioni.set(True)).grid(row=r, column=0, sticky="w"); r += 1
        ttk.Label(s3, text="Senza dettagli si ottiene solo l'elenco degli insegnamenti (molto più veloce).",
                  foreground="#666", wraplength=380).grid(row=r, column=0, sticky="w"); r += 1

        ttk.Label(s3, text="Avanzate", font=("", 9, "bold")).grid(row=r, column=0, sticky="w", pady=(8, 0)); r += 1
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        self.var_cache = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="Riusa le pagine già scaricate", variable=self.var_cache).grid(row=0, column=0)
        ttk.Button(f, text="Svuota cache", command=self._svuota_cache).grid(row=0, column=1, padx=8)
        f = ttk.Frame(s3); f.grid(row=r, column=0, sticky="w"); r += 1
        ttk.Label(f, text="Pausa tra le richieste (secondi)").grid(row=0, column=0)
        self.var_delay = tk.DoubleVar(value=0.4)
        ttk.Spinbox(f, from_=0.1, to=5, increment=0.1, width=5, textvariable=self.var_delay).grid(row=0, column=1, padx=6)
        ttk.Label(f, text="Richieste in parallelo").grid(row=0, column=2, padx=(10, 0))
        self.var_paralleli = tk.IntVar(value=ps.PARALLELI_DEFAULT)
        ttk.Spinbox(f, from_=1, to=8, increment=1, width=4, textvariable=self.var_paralleli).grid(row=0, column=3, padx=6)
        self.lbl_cache = ttk.Label(s3, foreground="#666")
        self.lbl_cache.grid(row=r, column=0, sticky="w"); r += 1
        self._aggiorna_info_cache()

        # ④ avvio
        s4 = Sezione(self, "④ Scarica")
        s4.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        s4.columnconfigure(1, weight=1)
        ttk.Label(s4, text="Salva in").grid(row=0, column=0, sticky="w")
        self.var_out = tk.StringVar()
        ttk.Entry(s4, textvariable=self.var_out).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(s4, text="Sfoglia…", command=self._scegli_out).grid(row=0, column=2)
        ttk.Label(s4, text="(vuoto = nome automatico nella cartella output)", foreground="#666").grid(row=0, column=3, padx=6)
        f = ttk.Frame(s4); f.grid(row=1, column=0, columnspan=4, sticky="ew", pady=6)
        f.columnconfigure(3, weight=1)
        self.btn_avvia = ttk.Button(f, text="▶  Avvia scaricamento", style="Accent.TButton", command=self._avvia)
        self.btn_avvia.grid(row=0, column=0, ipadx=10, ipady=3)
        self.btn_stop = ttk.Button(f, text="■  Interrompi", command=self._interrompi, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=8)
        self.lbl_prog = ttk.Label(f, text="", width=58)
        self.lbl_prog.grid(row=0, column=2, padx=8)
        self.pbar = ttk.Progressbar(f, mode="determinate")
        self.pbar.grid(row=0, column=3, sticky="ew")
        self.log = ScrolledText(s4, height=8, font=("Consolas" if sys.platform.startswith("win") else "Monospace", 9),
                                state="disabled", wrap="none")
        self.log.grid(row=2, column=0, columnspan=4, sticky="nsew")
        s4.rowconfigure(2, weight=1)

    # ---------------------------------------------------------------- catalogo
    def _in_thread(self, funzione, fine):
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
            self.lbl_stato.config(text="Sito non raggiungibile: controlla la connessione", foreground="#b00020")
            messagebox.showerror("Errore di connessione", f"Impossibile raggiungere il sito dei manifesti.\n\n{err}")
            return
        anni, sedi = res
        self.map_aa = {o["testo"]: o["valore"] for o in anni}
        self.map_sede = {o["testo"]: o["valore"] for o in sedi}
        self.cb_aa.config(values=list(self.map_aa))
        self.cb_sede.config(values=list(self.map_sede))
        self.cb_aa.set(VUOTO)
        self.cb_sede.set(VUOTO)
        self.lbl_stato.config(text="Scegli anno accademico e sede", foreground="#666")

    def _carica_corsi(self):
        aa, sede = self.map_aa.get(self.cb_aa.get()), self.map_sede.get(self.cb_sede.get())
        if not (aa and sede):
            return
        self.lbl_stato.config(text="Carico l'elenco dei corsi…", foreground="#666")
        self.catalogo = None
        self._riempi_albero()
        self.n_richiesta_corsi += 1
        n = self.n_richiesta_corsi
        self._in_thread(lambda: ps.catalogo(aa, sede), lambda cat, err: self._corsi_caricati(cat, err, n))

    def _corsi_caricati(self, cat, err, n_richiesta):
        if n_richiesta != self.n_richiesta_corsi:
            return  # nel frattempo l'utente ha cambiato anno o sede
        if err:
            self.lbl_stato.config(text="Errore nel caricare i corsi", foreground="#b00020")
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
        self.lbl_stato.config(text=f"{n} corsi disponibili", foreground="#1b7f3b")
        self._riempi_albero()

    # ---------------------------------------------------------------- albero
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
        n = len(self.selezionati)
        self.lbl_sel.config(text=f"{n} corsi selezionati" if n else "Nessun corso selezionato")

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

    def _seleziona_visibili(self, on):
        self.selezionati |= {c["codice"] for _, c in self._corsi_visibili()}
        self._aggiorna_spunte()

    def _deseleziona_tutti(self):
        self.selezionati.clear()
        self._aggiorna_spunte()

    # ---------------------------------------------------------------- opzioni
    def _anni_cambiati(self):
        if any(v.get() for v in self.var_anni.values()):
            self.var_anni_tutti.set(False)

    def _anni_tutti_cambiato(self):
        if self.var_anni_tutti.get():
            for v in self.var_anni.values():
                v.set(False)

    def _dettagli_cambiati(self):
        if not self.var_scaglioni.get():
            self.var_orari.set(False)

    def _aggiorna_info_cache(self):
        n = len(list(CACHE.glob("*.html"))) if CACHE.exists() else 0
        mb = sum(f.stat().st_size for f in CACHE.glob("*.html")) / 1e6 if n else 0
        self.lbl_cache.config(text=f"Cache: {n} pagine salvate ({mb:.0f} MB). Svuotala per avere dati aggiornati.")

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
        if not self.selezionati or not self.catalogo:
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
                         for c in s["corsi"] if c["codice"] in self.selezionati})
        return ps.Opzioni(
            aa=aa, sede=sede, scuole=scuole, corsi=sorted(self.selezionati),
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

    def _avvia(self):
        opt = self._opzioni()
        if not opt:
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
        ttk.Label(s1, text="File").grid(row=0, column=0)
        self.cb_file = ttk.Combobox(s1, state="readonly")
        self.cb_file.grid(row=0, column=1, sticky="ew", padx=6)
        self.cb_file.bind("<<ComboboxSelected>>", lambda e: self._carica_file())
        ttk.Button(s1, text="Sfoglia…", command=self._sfoglia).grid(row=0, column=2)
        ttk.Button(s1, text="Aggiorna elenco", command=self.aggiorna_elenco).grid(row=0, column=3, padx=4)
        ttk.Button(s1, text="Apri cartella output", command=lambda: apri_percorso(OUTPUT)).grid(row=0, column=4)
        self.lbl_file = ttk.Label(s1, foreground="#666")
        self.lbl_file.grid(row=1, column=0, columnspan=5, sticky="w", pady=(4, 0))

        s2 = Sezione(self, "② Tabella e filtri")
        s2.grid(row=1, column=0, sticky="ew", pady=8)
        f = ttk.Frame(s2)
        f.grid(row=0, column=0, columnspan=8, sticky="w")
        self.var_tab = tk.StringVar(value="")
        for i, (k, t) in enumerate(NOMI_TABELLE.items()):
            ttk.Radiobutton(f, text=t, value=k, variable=self.var_tab, command=self._tabella_cambiata).grid(row=0, column=i, padx=(0, 14))
        self.lbl_desc = ttk.Label(f, foreground="#666")
        self.lbl_desc.grid(row=0, column=len(NOMI_TABELLE), padx=8)

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
            self.cb_filtri[col], self.lbl_filtri[col] = cb, lbl
        for c in (1, 3, 5, 7):
            s2.columnconfigure(c, weight=1)
        f = ttk.Frame(s2)
        f.grid(row=3, column=0, columnspan=8, sticky="ew", pady=(4, 0))
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Cerca testo").grid(row=0, column=0)
        self.var_cerca = tk.StringVar()
        e = ttk.Entry(f, textvariable=self.var_cerca)
        e.grid(row=0, column=1, sticky="ew", padx=6)
        self._timer = None
        self.var_cerca.trace_add("write", lambda *a: self._rinvia_aggiornamento())
        self.var_unisci = tk.BooleanVar()
        ttk.Checkbutton(f, text="Una riga sola per ciò che si ripete in più corsi (unisce i doppioni)", variable=self.var_unisci,
                        command=self._aggiorna_anteprima).grid(row=0, column=2, padx=8)
        ttk.Button(f, text="Colonne…", command=self._scegli_colonne).grid(row=0, column=3)
        ttk.Button(f, text="Azzera filtri", command=self._azzera_filtri).grid(row=0, column=4, padx=(6, 0))

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
        for i, (fmt, testo) in enumerate([("xlsx", "Excel (.xlsx)"), ("csv", "CSV (.csv)"),
                                          ("html", "Pagina web (.html)"), ("json", "JSON (.json)")]):
            ttk.Button(s4, text=testo, command=lambda f=fmt: self._esporta(f)).grid(row=0, column=i, padx=(0, 6))
        ttk.Button(s4, text="Excel con tutte le tabelle", command=self._esporta_tutte).grid(row=0, column=4, padx=(12, 6))
        ttk.Separator(s4, orient="vertical").grid(row=0, column=5, sticky="ns", padx=10)
        ttk.Label(s4, text="Orario settimanale:").grid(row=0, column=6)
        self.cb_cal = ttk.Combobox(s4, state="readonly", width=46, values=list(ex.RAGGRUPPA.values()))
        self.cb_cal.set(VUOTO)
        self.cb_cal.grid(row=0, column=7, padx=6)
        ttk.Button(s4, text="Crea calendario", command=self._calendario).grid(row=0, column=8)

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
            self.tab = ex.tabelle(self.dati)
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
        w.transient(self)
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
        """Esegue scrivi(); se va bene propone di aprire il file, altrimenti spiega il problema."""
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            scrivi()
        except PermissionError:
            messagebox.showerror("File in uso", "Non riesco a scrivere il file: forse è aperto in Excel "
                                 "o in un altro programma? Chiudilo e riprova.")
            return
        except Exception as e:
            messagebox.showerror("Salvataggio non riuscito", f"{path}\n\n{e}")
            return
        finally:
            self.config(cursor="")
        if messagebox.askyesno("Salvato", f"{testo}\n\n{path}\n\nVuoi aprirlo adesso?"):
            apri_percorso(path)

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
        self.title("Manifesti degli Studi PoliMi – scaglioni e orari")
        self.chiusura_richiesta = False
        self.protocol("WM_DELETE_WINDOW", self._chiudi)
        self._stile()
        w, h = min(1400, self.winfo_screenwidth() - 60), min(900, self.winfo_screenheight() - 100)
        self.geometry(f"{w}x{h}+20+20")
        self.minsize(1000, 640)
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.scarica = SchedaScarica(self, self.nb)
        self.esplora = SchedaEsplora(self, self.nb)
        self.nb.add(self.scarica, text="  1 · Scarica dati dal sito  ")
        self.nb.add(self.esplora, text="  2 · Esplora ed esporta  ")
        if self.esplora.files:
            ttk.Label(self, text="Hai già dei dati scaricati: puoi passare direttamente alla scheda 2.",
                      foreground="#666").pack(anchor="w", padx=12, pady=(0, 6))

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
