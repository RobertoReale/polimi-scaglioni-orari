#!/usr/bin/env python3
"""
Ricerche al volo sul sito del Politecnico di Milano, senza scaricare corsi interi.

Ogni ricerca è una funzione che restituisce un Risultato: una o più tabelle (liste di dict, come
quelle di esporta.py) e qualche nota per l'utente. La usano sia il terminale (main) sia la
scheda «Cerca sul sito» dell'interfaccia.

LE RICERCHE E LE PAGINE DEL SITO CHE USANO
    insegnamenti(testo, docente)    insegnamenti con i loro docenti          elenco (*)
    docenti(nome)                   docenti per nome                         elenco (*)
    scheda_docente(codice o nome)   dati, insegnamenti, scaglioni e orario   scheda del docente
    chi_insegna(insegnamento, fasce) i docenti di un insegnamento con il loro orario, ordinati
                                    per quante fasce orarie richieste coprono    (*) + scheda del docente
    occupazione_aule(giorno, …)     chi occupa ogni aula e quando; aule libere   sito «Spazi»
    info_corso(corso, pagina)       struttura, elenco docenti, programmi interdisciplinari, scambi
    vecchi_ordinamenti(…)           insegnamenti degli ordinamenti precedenti al D.M. 509

    (*) Nonostante il nome, la pagina «Erogati in lingua Inglese» elenca tutti gli insegnamenti,
        ognuno con i suoi docenti e il loro codice (k_doc). È la fonte più comoda per sapere chi
        insegna cosa, e continua a funzionare quando la pagina di dettaglio dei manifesti è guasta.
        Cercando «%%%» la pagina restituisce l'elenco completo dell'anno: elenco() lo scarica una
        volta (15-20 secondi), lo tiene su disco per qualche ora e le ricerche lo filtrano in locale,
        tollerando parole come «e», «ed», «di» e gli accenti (corrisponde()).

ORARI APPROSSIMATIVI
    Le lezioni iniziano e finiscono al quarto d'ora (08:15, 10:15…). Ovunque si confrontano orari
    c'è una tolleranza di TOLLERANZA minuti: «dalle 16 alle 18» trova la lezione 16:15–18:15.

LA SCHEDA DEL DOCENTE (RicercaPerDocentiPublic.do)
    ?evn_prodotti=EVENTO&k_doc=…&aa=…                    dati del docente
    ?evn_DIDATTICA_AJAX=evento&k_doc=…&aa=…              i suoi insegnamenti (un blocco «tabs» per
                                                          ciascuno) con la tabella degli scaglioni
    ?evn_didattica_orario_incarico_AJAX=evento + qs      l'orario di un insegnamento: la stessa griglia
                                                          dei manifesti, letta da scarica_manifesti.parse_orario

Uso da terminale, esempi:
    python cerca.py chi-insegna "geometria e algebra lineare" --sede MI --fasce "gio 08:15-10:15, ven 10:15-13:15"
    python cerca.py insegnamenti "analisi matematica 1" --sede MI
    python cerca.py docente rossi
    python cerca.py aule --sede MIA --giorno 15/10/2026 --cerca geometria
    python cerca.py aule --sede MIA --giorno 15/10/2026 --libere --dalle 10:15 --alle 12:15
    python cerca.py corso --elenca
    python cerca.py corso 531 --mostra docenti
    python cerca.py vecchi-ordinamenti --insegnamento geometria
Ogni comando mostra le sue opzioni con --help. Con --out FILE.xlsx (o .csv, .html, .json) salva i risultati.
"""
import argparse
import importlib.util
import json
import os
import re
import shutil
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from bs4 import BeautifulSoup, Comment, NavigableString

import esporta as ex
import scarica_manifesti as sm
from scarica_manifesti import ErroreSito, Interrotto, soup, txt

CONTROLLER = sm.HOST + "/manifesti/manifesti/controller/"
URL_DOCENTI = CONTROLLER + "ricerche/RicercaPerDocentiPublic.do"
URL_INS_DOCENTI = CONTROLLER + "ricerche/RicercaInsegnamentiErogatiInLinguaInglesePublic.do"
URL_VO = CONTROLLER + "ricerche/RicercaPerInsegnamentoVOPublic.do"
URL_PER_INSEGNAMENTO = CONTROLLER + "ricerche/RicercaPerInsegnamentoPublic.do"
URL_SPAZI = sm.HOST + "/spazi/spazi/controller/OccupazioniGiornoEsatto.do"

PAGINE_CORSO = {  # chiave: (nome per l'utente, pagina del sito)
    "struttura": ("Struttura del corso", CONTROLLER + "MostraIndirizziPublic.do"),
    "docenti": ("Elenco docenti", CONTROLLER + "MostraFacultyPublic.do"),
    "interdisciplinari": ("Programmi interdisciplinari", CONTROLLER + "extra/ProgrammiInterdisciplinariPublic.do"),
    "scambi": ("Scambi internazionali", CONTROLLER + "extra/ScambiInternazionaliPublic.do"),
}

GIORNI_ABBR = {g[:3].lower(): g for g in sm.GIORNI}          # "gio" -> "Giovedì"
GIORNI_SPAZI = {"Lun": "Lunedì", "Mar": "Martedì", "Mer": "Mercoledì", "Gio": "Giovedì",
                "Ven": "Venerdì", "Sab": "Sabato", "Dom": "Domenica"}
MAX_GIORNI_AULE = 14
MAX_INSEGNAMENTI_COMPLETI = 3                        # chi_insegna: oltre, niente elenchi docenti dei corsi
TOLLERANZA = 15                                      # minuti, vedi «ORARI APPROSSIMATIVI»
ORE_LEZIONE = [f"{h:02d}:15" for h in range(8, 21)]  # 08:15 … 20:15, per i menu dell'interfaccia

# nomi leggibili delle colonne nuove (gli altri sono in esporta.LABELS)
ex.LABELS.update({
    "docente": "Docente", "codice_docente": "Cod. docente", "url_docente": "Scheda docente",
    "qualifica": "Qualifica", "ruolo": "Ruolo", "sede": "Sede", "iscritti": "Studenti iscritti",
    "insegnamenti": "Insegnamenti", "track": "Track (piano)", "voce": "Voce", "valore": "Valore",
    "corrispondenze": "Fasce coperte", "fasce": "Quali fasce", "data": "Data",
    "descrizione": "Descrizione sul sito", "libera": "Libera", "gruppo": "Gruppo", "paese": "Paese",
})


# ============================================================ risultato

@dataclass
class Tabella:
    nome: str
    colonne: list
    righe: list


@dataclass
class Risultato:
    titolo: str
    tabelle: list = field(default_factory=list)   # [Tabella]; la prima è la principale
    note: list = field(default_factory=list)      # frasi per l'utente (avvisi, suggerimenti)

    @property
    def n_righe(self):
        return sum(len(t.righe) for t in self.tabelle)


def _cli(log=None, stop=None, delay=0.3):
    """Client senza cache: le ricerche devono riflettere il sito di oggi."""
    return sm.Client(delay=delay, cache_dir=None, retries=3, log=log or (lambda m: None), stop=stop)


IN_LINEA = {"b", "strong", "i", "em", "u", "font", "span", "a", "sup", "sub", "small", "big"}


def _testo(el):
    """Testo di un elemento. I tag «in linea» (<b>, <span>, …) si uniscono senza spazi, perché il sito
    evidenzia la parola cercata anche dentro le parole (es. G<b>E</b>OMETRIA); gli altri tag
    (<div>, <br>, <td>…) separano con uno spazio."""
    if el is None:
        return ""
    pezzi = []
    for d in el.descendants:
        if isinstance(d, NavigableString):
            if not isinstance(d, Comment) and d.parent.name not in ("script", "style"):
                pezzi.append(str(d))
        elif d.name not in IN_LINEA:
            pezzi.append(" ")
    return re.sub(r"\s+", " ", "".join(pezzi)).strip()


def _testo_cella(td):
    """Testo di una cella; se è solo un'immagine (es. il pallino «anno attivo»), il suo testo alternativo."""
    t = _testo(td)
    if not t:
        img = td.find("img", alt=True)
        t = (img.get("alt") or "").strip() if img else ""
    return t


def _k_doc(href):
    m = re.search(r"k_doc=(\d+)", href or "")
    return m.group(1) if m else None


def url_docente(k_doc, aa):
    return f"{URL_DOCENTI}?evn_prodotti=EVENTO&k_doc={k_doc}&aa={aa}&lang=IT"


def _nome_proprio(maiuscolo):
    """'ROSSI MARIO' -> 'Rossi Mario' (l'occupazione aule scrive i nomi in maiuscolo)."""
    return " ".join(p.capitalize() if p.isupper() else p for p in maiuscolo.split()) if maiuscolo else ""


def _hhmm(minuti):
    return f"{minuti // 60:02d}:{minuti % 60:02d}"


def _minuti(hhmm):
    """'08:15', '8.15' o '8' -> minuti dalla mezzanotte."""
    m = re.fullmatch(r"\s*(\d{1,2})(?:[:.h](\d{2}))?\s*", hhmm or "")
    if not m or int(m.group(1)) > 24 or int(m.group(2) or 0) > 59:
        raise ValueError(f"Ora non valida: «{hhmm}». Scrivila come 08:15 (oppure solo 8).")
    return int(m.group(1)) * 60 + int(m.group(2) or 0)


# ============================================================ confronto di nomi

# parole che non servono a distinguere un nome: «geometria ed algebra» = «geometria e algebra»
PAROLE_VUOTE = set("""e ed o di d del dell dello della dei degli delle da dal dall dalla dai a ad al all allo
    alla ai agli alle il lo la l i gli le un una uno per con in nel nell nella nei su sul sull sulla tra fra
    and of the for to on at an""".split())


def _normale(testo):
    """Minuscolo e senza accenti: «Città» -> «citta»."""
    s = unicodedata.normalize("NFKD", str(testo or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def parole(testo):
    """Le parole significative di un testo, normalizzate."""
    return [p for p in re.findall(r"[a-z0-9]+", _normale(testo)) if p not in PAROLE_VUOTE]


def corrisponde(cercato, *campi):
    """Vero se ogni parola cercata è l'inizio di una parola dei campi, in qualsiasi ordine.
    «geometria ed algebra lineare», «geom alg» e «082747» trovano tutti GEOMETRIA E ALGEBRA LINEARE."""
    cercate = parole(cercato)
    bersaglio = re.findall(r"[a-z0-9]+", _normale(" ".join(c for c in campi if c)))
    return bool(cercate) and all(any(b.startswith(p) for b in bersaglio) for p in cercate)


# ============================================================ scelte disponibili

_scelte = {}


def scelte(log=None):
    """Anni accademici e sedi dei manifesti, sedi del sito Spazi: [(codice, nome)], letti dal sito."""
    if not _scelte:
        cli = _cli(log)
        pg = soup(cli.get(URL_INS_DOCENTI, {"evn_default": "EVENTO", "lang": "IT"}))
        _scelte["aa"] = [(o["valore"], o["testo"]) for o in sm.select_options(pg, "aa")]
        _scelte["aa_attuale"] = next((o["valore"] for o in sm.select_options(pg, "aa") if o["selezionata"]),
                                     _scelte["aa"][0][0] if _scelte["aa"] else None)
        _scelte["sedi"] = [(o["valore"], o["testo"]) for o in sm.select_options(pg, "sede")
                           if o["valore"] != "ALL_SEDI"]
        pg = soup(cli.get(URL_SPAZI, {"evn_init": "event", "jaf_currentWFID": "main"}))
        _scelte["sedi_aule"] = [(o["valore"], o["testo"]) for o in sm.select_options(pg, "csic")
                                if o["valore"] != "tutte"]
    return _scelte


def _aa(aa):
    return str(aa) if aa else scelte()["aa_attuale"]


def _nome_sede(sede):
    """'MI' -> 'Milano Leonardo' (come la scrive la scheda del docente)."""
    nome = dict(scelte()["sedi"]).get(sede, "")
    return re.sub(r"\s*\(\w+\)$", "", nome)


# ============================================================ tabelle generiche

def _e_intestazione(cell):
    """Una cella d'intestazione: <th>, classe HeadColumn, oppure testo tutto in grassetto."""
    if cell.name == "th" or any("HeadColumn" in c for c in cell.get("class") or []):
        return True
    t = txt(cell)
    if not t:
        return False
    for b in cell.find_all(["b", "strong", "div", "span"]):
        if (b.name in ("b", "strong") or "bold" in (b.get("style") or "")) and txt(b) == t:
            return True
    return False


def righe_tabella(table, intestazioni=None):
    """Le righe di una tabella del sito come dict {intestazione: testo}.
    - più righe d'intestazione consecutive vengono combinate (es. «Anni di corso 1°»);
    - una riga con una sola cella larga quanto la tabella (es. «Corso di Studi …») diventa il campo
      «gruppo» delle righe che seguono;
    - le righe «voce | valore» (come le schede informative) diventano {"voce", "valore"}.
    Ogni riga ha anche "_celle": {intestazione: cella html}, per leggere i link."""
    grid = sm.table_grid(table)
    larghezza = max((c + s for row in grid for c, s, _, _ in row), default=0)
    nomi, gruppo, righe, blocco = dict(intestazioni or {}), None, [], []
    for row in grid + [None]:
        if row is not None and row and all(_e_intestazione(cell) for _, _, cell, _ in row) \
                and not (len(row) == 1 and larghezza > 1):
            blocco.append(row)
            continue
        if blocco:  # fine di un gruppo di righe d'intestazione
            nomi = sm.header_names(blocco, len(blocco))
            blocco = []
        if not row:
            continue
        if len(row) == 1 and (row[0][1] >= larghezza or larghezza == 1):
            gruppo = _testo_cella(row[0][2]) or gruppo
            continue
        classi = ["1" if any(c.lower().startswith("elementinfocard1") for c in cell.get("class") or [])
                  else "2" for _, _, cell, _ in row]
        if not nomi and len(row) % 2 == 0 and classi == ["1", "2"] * (len(row) // 2):
            for k in range(0, len(row), 2):
                voce, valore = _testo_cella(row[k][2]), _testo_cella(row[k + 1][2])
                if voce:
                    righe.append({"gruppo": gruppo, "voce": voce, "valore": valore,
                                  "_celle": {"valore": row[k + 1][2]}})
            continue
        r = {"gruppo": gruppo, "_celle": {}}
        for col, _, cell, _ in row:
            nome = nomi.get(col) or f"colonna {col + 1}"
            r[nome] = _testo_cella(cell)
            r["_celle"][nome] = cell
        if any(v for k, v in r.items() if k not in ("gruppo", "_celle")):
            righe.append(r)
    return righe


def _pulisci(righe):
    """Toglie i campi interni ("_celle") e la colonna "gruppo" se ha sempre lo stesso valore.
    Restituisce (colonne, righe, gruppo unico o None)."""
    colonne = []
    for r in righe:
        for k in r:
            if not k.startswith("_") and k not in colonne:
                colonne.append(k)
    gruppi = {r.get("gruppo") for r in righe}
    unico = next(iter(gruppi)) if len(gruppi) == 1 else None
    if len(gruppi) <= 1 and "gruppo" in colonne:
        colonne.remove("gruppo")
    return colonne, [{c: r.get(c) for c in colonne} for r in righe], unico


def _blocchi_pagina(page):
    """Le tabelle di dati di una pagina informativa del sito (escluso il menu e il modulo di scelta),
    ognuna col titolo che la precede. Tabelle consecutive con le stesse colonne (es. un Paese
    dopo l'altro) diventano una sola, col titolo nella colonna «gruppo»."""
    blocchi, titolo = [], None
    for t in page.find_all("table"):
        if t.find_parent("form") or "TableCommand" in (t.get("class") or []) or t.get("id") == "poliheader":
            continue
        if t.find_parent("table", class_=lambda c: c and ("TableDati" in c or "ttree" in c)):
            continue
        celle = [td for tr in sm.own_rows(t) for td in tr.find_all(["td", "th"], recursive=False)]
        if len(celle) == 1:
            classi = " ".join(celle[0].get("class") or [])
            if "TitleInfoCard" in classi or "SimpleCard" in classi:
                titolo = _testo_cella(celle[0]) or titolo
            continue
        if not (set(t.get("class") or []) & {"BoxInfoCard", "TableDati", "ttree"}):
            continue
        if titolo == "Legenda" or t.find_parent(id=re.compile("legenda", re.I)):
            continue  # la legenda dei simboli, nella colonna di sinistra
        righe = righe_tabella(t)
        if not righe:
            continue
        for r in righe:
            r["gruppo"] = r.get("gruppo") or titolo
        chiavi = [k for k in righe[0] if k not in ("gruppo", "_celle")]
        if blocchi and blocchi[-1][1] == chiavi:
            blocchi[-1][2].extend(righe)
        else:
            blocchi.append([titolo, chiavi, righe])
    return blocchi


def _messaggi_modulo(page):
    """Le frasi informative dentro il modulo di scelta (es. «Il corso di studi non offre
    programmi interdisciplinari»): righe «voce | testo» senza menu né caselle."""
    out = []
    for form in page.find_all("form"):
        for tr in form.find_all("tr"):
            tds = tr.find_all("td", recursive=False)
            if len(tds) == 2 and not tr.find(["select", "input"]) and txt(tds[1]):
                out.append(f"{txt(tds[0])}: {txt(tds[1])}")
    return out


# ============================================================ insegnamenti e docenti

TUTTI = "%%%"                                    # la pagina (*) vuole 3 caratteri: «%%%» trova tutto
CARTELLA_ELENCHI = sm.HERE / "cache" / "ricerche"
VALIDITA_ELENCO = 12 * 3600                      # secondi: gli insegnamenti cambiano di rado
# lxml è facoltativo: legge l'elenco completo (4 MB) in metà tempo
PARSER_VELOCE = "lxml" if importlib.util.find_spec("lxml") else "html.parser"


def _insegnamenti_docenti(cli, testo="", docente="", aa=None, sede=None):
    """Righe della pagina (*): una per insegnamento e docente."""
    aa = _aa(aa)
    cli.get(URL_INS_DOCENTI, {"evn_default": "EVENTO", "aa": aa, "lang": "IT"})  # apre la sessione
    html = cli.post(URL_INS_DOCENTI, {
        "lang": "IT", "aa": aa, "sede": sede or "ALL_SEDI", "k_cf": "-1", "k_corso_la": "-1",
        "aree": "-1", "codDescr": testo or "", "n_docente": docente or "", "evn_ricercainschinc": "Aggiorna"})
    table = BeautifulSoup(html, PARSER_VELOCE).find("table", class_="TableDati")
    out = []
    if not table:
        return out
    corso = None
    for tr in table.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) == 1:
            corso = _testo(tds[0]).replace("Corso di Studi ", "", 1)
            continue
        if len(tds) < 6:
            continue
        base = {"corso": corso, "codice": _testo(tds[2]), "insegnamento": _testo(tds[3]),
                "sede": _testo(tds[1]), "periodo": _testo(tds[5]), "tipo": _testo(tds[0])}
        link = tds[4].find_all("a", href=True) or [None]
        for a in link:
            k = _k_doc(a["href"]) if a else None
            out.append({**base, "docente": _testo(a) if a else _testo(tds[4]), "codice_docente": k,
                        "url_docente": url_docente(k, aa) if k else None})
    return out


_elenchi = {}                      # anno -> righe di elenco()
_elenchi_lock = threading.Lock()   # l'interfaccia lo carica in un thread mentre si può già cercare


def elenco(aa=None, log=None, stop=None):
    """Tutti gli insegnamenti dell'anno con i loro docenti (una riga per coppia, come
    _insegnamenti_docenti). Il sito impiega 15-20 secondi: si tiene in memoria e su disco."""
    aa = _aa(aa)
    if _elenchi.get(aa):  # già pronto: senza aspettare chi sta scaricando l'elenco di un altro anno
        return _elenchi[aa]
    with _elenchi_lock:
        if not _elenchi.get(aa):
            _elenchi[aa] = _leggi_elenco(aa, log, stop)
        return _elenchi[aa]


def elenco_pronto(aa=None):
    """Vero se elenco(aa) risponde subito (già in memoria)."""
    return bool(_elenchi.get(_aa(aa)))


def _leggi_elenco(aa, log, stop):
    f = CARTELLA_ELENCHI / f"insegnamenti_{aa}.json"
    try:
        if time.time() - f.stat().st_mtime < VALIDITA_ELENCO:
            return json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass  # niente copia recente: si scarica
    (log or (lambda m: None))("Scarico l'elenco di tutti gli insegnamenti (una volta ogni qualche ora)…")
    righe = _insegnamenti_docenti(_cli(log, stop), TUTTI, aa=aa)
    if not righe:
        raise ErroreSito("l'elenco degli insegnamenti è vuoto")
    f.parent.mkdir(parents=True, exist_ok=True)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(righe, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, f)  # scrittura atomica: mai un file a metà
    return righe


def _filtra(righe, testo="", docente="", sede=None):
    """Le righe con l'insegnamento `testo` (nome o codice), il docente `docente`, nella sede `sede` (codice)."""
    nome_sede = _nome_sede(sede) if sede else None
    return [r for r in righe
            if (not testo or corrisponde(testo, r["codice"], r["insegnamento"]))
            and (not docente or corrisponde(docente, r["docente"], r["codice_docente"]))
            and (not nome_sede or r["sede"] == nome_sede)]


def nomi_insegnamenti(aa=None, sede=None):
    """[(«codice nome», codice)] degli insegnamenti dell'anno (della sede, se data), per i suggerimenti."""
    visti = {r["codice"]: r["insegnamento"] for r in _filtra(elenco(aa), sede=sede)}
    return sorted(((f"{c} {n}", c) for c, n in visti.items()), key=lambda v: _normale(v[0][7:]))


def nomi_docenti(aa=None):
    """[(nome, codice)] di tutti i docenti dell'anno, per i suggerimenti."""
    visti = {r["codice_docente"]: r["docente"] for r in elenco(aa) if r["codice_docente"]}
    return sorted(((n, c) for c, n in visti.items()), key=lambda v: _normale(v[0]))


def suggerimenti(voci, testo, massimo=30):
    """Le voci [(testo, valore)] che corrispondono a `testo`: prima quelle in cui la prima parola
    cercata viene presto («geometria» -> GEOMETRIA E ALGEBRA… prima di ANALISI… E GEOMETRIA)."""
    cercate = parole(testo)
    if not cercate:
        return []

    def posizione(voce):
        nomi = re.findall(r"[a-z0-9]+", _normale(voce[0]))
        return next((i for i, n in enumerate(nomi) if n.startswith(cercate[0])), len(nomi))
    return sorted((v for v in voci if corrisponde(testo, v[0])), key=posizione)[:massimo]


def insegnamenti(testo="", docente="", aa=None, sede=None, log=None, stop=None):
    """Insegnamenti con `testo` nel nome o nel codice (e/o con un docente `docente`), ognuno con i suoi docenti."""
    if not parole(testo) and not parole(docente):
        raise ValueError("Scrivi il nome o il codice dell'insegnamento, oppure il nome del docente.")
    righe = _filtra(elenco(aa, log, stop), testo, docente, sede)
    colonne = ["insegnamento", "codice", "docente", "sede", "periodo", "corso", "tipo",
               "codice_docente", "url_docente"]
    ris = Risultato(f"Insegnamenti «{testo or ''}»" + (f", docente «{docente}»" if docente else ""),
                    [Tabella("Insegnamenti e docenti", colonne, righe)])
    if not righe:
        ris.note.append("Nessun insegnamento trovato. Prova con meno parole o con il codice (es. 082747).")
    else:
        ris.note.append("Ogni insegnamento compare sotto uno solo dei corsi che lo hanno nel piano.")
    return ris


def docenti(nome, aa=None, log=None, stop=None):
    """Docenti il cui nome corrisponde a `nome`, ognuno con i suoi insegnamenti dell'anno.
    (La pagina «Cerca Docenti» del sito non serve: l'elenco degli insegnamenti ha già tutti i docenti.)"""
    if not parole(nome):
        raise ValueError("Scrivi il cognome (o una parte del nome) del docente.")
    aa = _aa(aa)
    per_doc = {}
    for r in _filtra(elenco(aa, log, stop), docente=nome):
        d = per_doc.setdefault(r["codice_docente"] or r["docente"], {
            "docente": r["docente"], "codice_docente": r["codice_docente"], "insegnamenti": [],
            "url_docente": r["url_docente"]})
        voce = f"{r['codice']} {r['insegnamento']} ({r['sede']})"
        if voce not in d["insegnamenti"]:
            d["insegnamenti"].append(voce)
    righe = sorted(({**d, "insegnamenti": " · ".join(d["insegnamenti"])} for d in per_doc.values()),
                   key=lambda r: _normale(r["docente"]))
    ris = Risultato(f"Docenti «{nome}»", [Tabella("Docenti", ["docente", "codice_docente", "insegnamenti",
                                                                "url_docente"], righe)])
    if not righe:
        ris.note.append("Nessun docente trovato tra chi insegna quest'anno.")
    elif len(righe) > 1:
        ris.note.append("Per la scheda di un docente scegli la sua riga (doppio clic) o usa il suo codice.")
    return ris


def _incarichi(cli, k_doc, aa):
    """Gli insegnamenti di un docente nell'anno: dati, scaglioni e il «qs» per chiederne l'orario."""
    page = soup(cli.get(URL_DOCENTI, {"evn_DIDATTICA_AJAX": "evento", "aa": aa, "k_doc": k_doc,
                                      "lang": "IT", "tab_ricerca": "1"}))
    out = []
    for tabs in page.find_all("div", class_="tabs"):
        tid = tabs.get("id")
        testa = tabs.find("table", class_="BoxInfoCard")
        info = (righe_tabella(testa) or [{}])[0] if testa else {}
        denom = info.get("Denominazione Insegnamento", "")
        codice, _, nome = denom.partition(" - ")
        lingua = None
        if testa and testa.find("img", src=re.compile(r"flags/")):
            lingua = re.search(r"flags/(\w+)\.", testa.find("img", src=re.compile(r"flags/"))["src"]).group(1)
        li_or = tabs.find("li", id=f"{tid}_tab_orario")
        li_sc = tabs.find("li", id=f"{tid}_tab_dettaglio")
        contenuto = page.find(id=f"{tid}_contenuto")
        if (contenuto is None or not contenuto.find("table")) and li_sc is not None and li_sc.get("qs"):
            contenuto = soup(cli.get(URL_DOCENTI + "?evn_didattica_dettaglio_incarico_AJAX=evento" + li_sc["qs"]))
        scaglioni = []
        tab_sc = contenuto.find("table", class_="TableDati") if contenuto else None
        for r in righe_tabella(tab_sc) if tab_sc else []:
            scaglioni.append({"corso": r.get("Corso di Studi"), "track": r.get("Track"),
                              "da": r.get("Da (compreso)"), "a": r.get("A (escluso)"),
                              "periodo_scaglione": r.get("Periodo")})
        out.append({
            "codice": codice.strip(), "insegnamento": nome.strip() or denom, "cfu": info.get("CFU"),
            "periodo": info.get("Periodo"), "sede": info.get("Sede"), "ruolo": info.get("Ruolo"),
            "iscritti": info.get("Studenti iscritti"), "lingua": lingua, "scaglioni": scaglioni,
            "qs_orario": li_or.get("qs") if li_or is not None else None,
            "orario_disattivato": li_or is not None and "ui-state-disabled" in (li_or.get("class") or []),
        })
    return out


def _orario_incarico(cli, inc):
    """Le lezioni settimanali di un insegnamento del docente ([] se l'orario non è pubblicato)."""
    if not inc["qs_orario"] or inc["orario_disattivato"]:
        return []
    frag = soup(cli.get(URL_DOCENTI + "?evn_didattica_orario_incarico_AJAX=evento" + inc["qs_orario"]))
    return sm.parse_orario(frag)


def _info_docente(cli, k_doc, aa):
    """Le voci della scheda del docente (Docente, Qualifica, Dipartimento, E-mail, …)."""
    page = soup(cli.get(URL_DOCENTI, {"evn_prodotti": "EVENTO", "k_doc": k_doc, "aa": aa, "lang": "IT"}))
    voci, fine = [], page.find(id="tabs_attivita")
    for t in page.find_all("table", class_="BoxInfoCard"):
        if fine is not None and t.find_previous(id="tabs_attivita") is not None:
            break
        if t.find_parent(id=re.compile("legenda")) or t.find_parent("table", class_="BoxInfoCard"):
            continue
        for tr in sm.own_rows(t):
            tds = tr.find_all("td", recursive=False)
            if len(tds) < 2 or "ElementInfoCard1" not in (tds[0].get("class") or []):
                continue
            interna = tds[1].find("table")
            if interna:  # es. l'orario di ricevimento: una tabellina; tengo solo le celle con un valore
                valore = " · ".join(
                    f"{k}: {v}" if k != "Note" else v
                    for r in righe_tabella(interna) for k, v in r.items()
                    if not k.startswith("_") and k != "gruppo" and v and v != "---" and not re.fullmatch(
                        r"Dalle\s*:\s*Alle\s*:", v))
            else:
                valore = txt(tds[1])
            if txt(tds[0]) and valore not in ("", "---"):
                voci.append({"voce": txt(tds[0]), "valore": valore})
    return voci


def _lezione(lz, **extra):
    return {**extra, "giorno": lz.get("giorno"), "inizio": lz.get("inizio"), "fine": lz.get("fine"),
            "aula": lz.get("aula"), "edificio": lz.get("aula_descrizione"), "attivita": lz.get("attivita"),
            "dal": lz.get("dal"), "al": lz.get("al")}


def _ordine_lezione(r):
    return (sm.GIORNI.index(r["giorno"]) if r.get("giorno") in sm.GIORNI else 9, r.get("inizio") or "")


def scheda_docente(chi, aa=None, log=None, stop=None):
    """La scheda di un docente (codice numerico, oppure nome se ne trova uno solo): dati,
    insegnamenti con i loro scaglioni e l'orario delle lezioni."""
    aa = _aa(aa)
    chi = str(chi).strip()
    if not chi.isdigit():
        trovati = docenti(chi, aa, log, stop)
        righe = trovati.tabelle[0].righe
        uguali = [r for r in righe if _normale(r["docente"]) == _normale(chi)]
        if len(uguali) == 1:  # «Rossi Mario» non è ambiguo anche se esiste «Rossi Mariolina»
            righe = uguali
        if len(righe) != 1 or not righe[0]["codice_docente"]:
            trovati.note.insert(0, "Ci sono più docenti con questo nome: scegline uno." if righe
                                else "Nessun docente con questo nome.")
            return trovati
        chi = righe[0]["codice_docente"]
    cli = _cli(log, stop)
    voci = _info_docente(cli, chi, aa)
    nome = next((v["valore"] for v in voci if v["voce"] == "Docente"), f"docente {chi}")
    incarichi = _incarichi(cli, chi, aa)
    righe_sc, righe_or = [], []
    for inc in incarichi:
        base = {k: inc[k] for k in ("codice", "insegnamento", "cfu", "periodo", "sede", "ruolo", "iscritti")}
        for sc in inc["scaglioni"] or [{}]:
            righe_sc.append({**base, **sc})
        for lz in _orario_incarico(cli, inc):
            righe_or.append(_lezione(lz, codice=inc["codice"], insegnamento=inc["insegnamento"]))
    righe_or.sort(key=_ordine_lezione)
    voci.append({"voce": "Pagina sul sito", "valore": url_docente(chi, aa)})
    ris = Risultato(f"{nome} – anno accademico {aa}/{int(aa) + 1}", [
        Tabella("Orario", ["giorno", "inizio", "fine", "aula", "insegnamento", "codice", "dal", "al",
                           "edificio", "attivita"], righe_or),
        Tabella("Insegnamenti e scaglioni", ["codice", "insegnamento", "corso", "track", "da", "a", "cfu",
                                             "periodo", "sede", "ruolo", "iscritti"], righe_sc),
        Tabella("Docente", ["voce", "valore"], voci)])
    if not incarichi:
        ris.note.append("Nessun insegnamento in quest'anno accademico.")
    elif not righe_or:
        ris.note.append("L'orario delle lezioni non è pubblicato.")
    return ris


# ============================================================ chi insegna (fasce orarie)

@dataclass
class Fascia:
    giorno: str          # "Giovedì"
    inizio: int = None   # minuti dalla mezzanotte; None: tutto il giorno («mar» = ha lezione il martedì)
    fine: int = None     # None: un istante («gio 08:15» = a lezione alle 08:15)

    def __str__(self):
        g = self.giorno[:3]
        if self.inizio is None:
            return g
        return f"{g} {_hhmm(self.inizio)}" + (f"–{_hhmm(self.fine)}" if self.fine is not None else "")

    def coperta_da(self, lz):
        """La lezione `lz` copre la fascia: stesso giorno e (intervallo) lo contiene tutto,
        oppure (istante) è in corso in quel momento, oppure (solo il giorno) è in quel giorno.
        Con TOLLERANZA minuti di margine, così «gio 16-18» trova la lezione 16:15–18:15."""
        if lz.get("giorno") != self.giorno or not lz.get("inizio") or not lz.get("fine"):
            return False
        if self.inizio is None:
            return True
        a, b = _minuti(lz["inizio"]), _minuti(lz["fine"])
        if self.fine is None:
            return a - TOLLERANZA <= self.inizio < b
        return a <= self.inizio + TOLLERANZA and b >= self.fine - TOLLERANZA


def leggi_fasce(testo):
    """«gio 08:15-10:15, ven 10:15, mar» -> [Fascia]. Giorni: lun mar mer gio ven sab (o il nome intero);
    le ore anche senza minuti («gio 8-10»); il giorno da solo vuol dire «a qualunque ora»."""
    fasce = []
    for pezzo in re.split(r"[,;\n]+", testo or ""):
        pezzo = pezzo.strip()
        if not pezzo:
            continue
        m = re.fullmatch(r"([a-zàèéìòù]+)(?:\s+([\d:.h]+)\s*(?:[-–]\s*([\d:.h]+))?)?", pezzo.lower())
        giorno = GIORNI_ABBR.get(m.group(1)[:3]) if m else None
        if not giorno:
            raise ValueError(f"Fascia non valida: «{pezzo}». Scrivila come «gio 08:15-10:15», «ven 10:15» "
                             "oppure solo «mar».")
        inizio = _minuti(m.group(2)) if m.group(2) else None
        fine = _minuti(m.group(3)) if m.group(3) else None
        if fine is not None and fine <= inizio:
            raise ValueError(f"Fascia non valida: «{pezzo}»: la fine viene prima dell'inizio.")
        fasce.append(Fascia(giorno, inizio, fine))
    return fasce


def _orario_breve(lezioni):
    return " · ".join(f"{(l['giorno'] or '?')[:3]} {l['inizio']}–{l['fine']} {l.get('aula') or ''}".strip()
                      for l in sorted(lezioni, key=_ordine_lezione))


def _scaglioni_brevi(scaglioni):
    gruppi = {}
    for s in scaglioni:
        if s.get("da") or s.get("a"):
            gruppi.setdefault(f"{s.get('da') or 'A'} – {s.get('a') or 'ZZZZ'}", []).append(s.get("corso"))
    return " · ".join(f"{k} ({len(v)} corsi/piani)" if len(v) > 1 else f"{k} ({v[0]})" for k, v in gruppi.items())


def _corsi_con_insegnamento(cli, codice, aa, sede=None):
    """I codici dei corsi di studio che hanno l'insegnamento `codice` nel piano («Ricerca per insegnamento»)."""
    cli.get(URL_PER_INSEGNAMENTO, {"evn_default": "EVENTO", "aa": aa, "lang": "IT"})  # apre la sessione
    page = soup(cli.post(URL_PER_INSEGNAMENTO + "?jaf_currentWFID=main", {
        "aa": aa, "k_cf": "-1", "sede": sede or "ALL_SEDI", "tipoCorso": "ALL_TIPO_CORSO", "ac_ins": "0",
        "semestre": "ALL_SEMESTRI", "aree": "-1", "tipoInsegnamento": "ALL_TIPO_INSEGNAMENTO",
        "insegn_ricerca": codice, "evn_default": "Esegui Ricerca", "lang": "IT"}))
    corsi = []
    for td in page.find_all("td"):  # «Corso di Studi Ing. Ind-Inf (1 liv.)(ord. 96/23) - MI (531) Ingegneria …»
        m = None if td.find("td") else re.match(r"Corso di Studi .*\((\d+)\)[^()]*$", txt(td))
        if m and m.group(1) not in corsi:
            corsi.append(m.group(1))
    return corsi


_docenti_corsi = {}  # (anno, corso) -> {(codice insegnamento, codice docente): nome}


def _docenti_del_corso(corso, aa, log=None, stop=None):
    """Chi insegna cosa secondo l'«Elenco docenti» del corso: {(codice insegnamento, codice docente): nome}."""
    if (aa, corso) not in _docenti_corsi:
        out = {}
        for t in info_corso(corso, "docenti", aa, log, stop).tabelle:
            for r in t.righe:
                codice = (r.get("Denominazione insegnamento") or "").split(" - ")[0].strip()
                if codice and r.get("codice_docente"):
                    out[(codice, r["codice_docente"])] = r.get("Docente")
        _docenti_corsi[(aa, corso)] = out
    return _docenti_corsi[(aa, corso)]


def _altri_docenti(cli, codici, aa, sede, log, stop):
    """{codice docente: nome} di chi insegna `codici` secondo l'elenco docenti dei corsi che li hanno nel piano.
    Servono perché l'elenco (*) mette ogni insegnamento sotto un solo corso, a volte senza i docenti degli
    altri corsi. Restituisce anche quanti corsi sono stati letti."""
    corsi = sorted({c for codice in codici for c in _corsi_con_insegnamento(cli, codice, aa, sede)})
    out = {}
    for _, corso, res, err in sm.in_parallelo(lambda c: _docenti_del_corso(c, aa, log, stop), corsi, 3):
        if err:
            (log or (lambda m: None))(f"Elenco docenti del corso {corso} non letto: {err}")
        for (codice, k_doc), nome in (res or {}).items():
            if codice in codici:
                out.setdefault(k_doc, nome)
    return out, len(corsi)


def chi_insegna(insegnamento, fasce="", aa=None, sede=None, log=None, stop=None, avanzamento=None):
    """I docenti di un insegnamento con il loro orario. Con `fasce` (es. «gio 08:15-10:15, ven 10:15-13:15»)
    li ordina per quante fasce coprono: in cima chi le copre tutte."""
    fasce = leggi_fasce(fasce) if isinstance(fasce, str) else list(fasce or [])
    if not parole(insegnamento):
        raise ValueError("Scrivi il nome o il codice dell'insegnamento.")
    aa = _aa(aa)
    log = log or (lambda m: None)
    avanzamento = avanzamento or (lambda fatti, totale: None)
    trovati = _filtra(elenco(aa, log, stop), testo=insegnamento, sede=sede)
    cli = _cli(log, stop)
    titolo = f"Chi insegna «{insegnamento}»" + (f" ({_nome_sede(sede)})" if sede else "")
    if not trovati:
        return Risultato(titolo, [], ["Nessun insegnamento trovato con questo nome o codice."])
    codici = {r["codice"]: r["insegnamento"] for r in trovati}
    per_doc = {r["codice_docente"]: r["docente"] for r in trovati if r["codice_docente"]}
    note = []
    if len(codici) <= MAX_INSEGNAMENTI_COMPLETI:
        log("Cerco i docenti anche negli elenchi dei corsi che hanno l'insegnamento…")
        altri, n_corsi = _altri_docenti(cli, codici, aa, sede, log, stop)
        nuovi = {k: v for k, v in altri.items() if k not in per_doc}
        per_doc.update(nuovi)
        if n_corsi:
            note.append(f"Docenti presi dall'elenco degli insegnamenti e dagli elenchi docenti di {n_corsi} corsi.")
    sede_nome = _nome_sede(sede) if sede else None
    log(f"{len(per_doc)} docenti da controllare")

    def leggi(k_doc):
        incs = [i for i in _incarichi(cli, k_doc, aa)
                if i["codice"] in codici and (not sede_nome or not i["sede"] or i["sede"] == sede_nome)]
        return [(inc, _orario_incarico(cli, inc)) for inc in incs]

    righe, lezioni, errori = [], [], []
    for n, (_, k_doc, res, err) in enumerate(sm.in_parallelo(leggi, list(per_doc), 3), 1):
        avanzamento(n, len(per_doc))
        nome = per_doc[k_doc]
        if err:
            errori.append(f"{nome}: {err}")
            continue
        for inc, orario in res:
            coperte = [f for f in fasce if any(f.coperta_da(lz) for lz in orario)]
            righe.append({
                "docente": nome, "corrispondenze": f"{len(coperte)}/{len(fasce)}" if fasce else None,
                "fasce": ", ".join(map(str, coperte)) or None, "orario": _orario_breve(orario) or "non pubblicato",
                "scaglioni": _scaglioni_brevi(inc["scaglioni"]), "insegnamento": inc["insegnamento"],
                "codice": inc["codice"], "sede": inc["sede"], "codice_docente": k_doc,
                "url_docente": url_docente(k_doc, aa), "_n": len(coperte)})
            lezioni += [_lezione(lz, docente=nome, codice=inc["codice"], insegnamento=inc["insegnamento"])
                        for lz in orario]
    righe.sort(key=lambda r: (-r["_n"], r["docente"].lower()))
    if len(codici) > 1:
        note.append(f"«{insegnamento}» corrisponde a {len(codici)} insegnamenti: "
                    + "; ".join(f"{c} {n}" for c, n in codici.items())
                    + ". Per restringere scrivi il codice"
                    + ("." if len(codici) <= MAX_INSEGNAMENTI_COMPLETI else
                       ", così si leggono anche gli elenchi docenti dei corsi e l'elenco è completo."))
    if fasce:
        tutte = [r for r in righe if r["_n"] == len(fasce)]
        if tutte:
            note.insert(0, "Copre tutte le fasce: " + ", ".join(f"{r['docente']} ({r['codice']})" for r in tutte))
        else:
            note.insert(0, "Nessun docente copre tutte le fasce; in cima chi ne copre di più.")
    note += [f"Non letto: {e}" for e in errori]
    for r in righe:
        del r["_n"]
    lezioni.sort(key=lambda r: (r["docente"].lower(),) + _ordine_lezione(r))
    colonne = (["docente", "corrispondenze", "fasce"] if fasce else ["docente"]) + [
        "orario", "scaglioni", "insegnamento", "codice", "sede", "codice_docente", "url_docente"]
    return Risultato(titolo + (f" – fasce: {', '.join(map(str, fasce))}" if fasce else ""), [
        Tabella("Docenti", colonne, righe),
        Tabella("Orario dei docenti", ["docente", "giorno", "inizio", "fine", "aula", "insegnamento", "codice",
                                       "dal", "al", "edificio", "attivita"], lezioni)], note)


# ============================================================ occupazione aule (sito Spazi)

def _data(testo):
    try:
        return datetime.strptime(testo.strip(), "%d/%m/%Y").date()
    except (ValueError, AttributeError):
        raise ValueError(f"Data non valida: «{testo}». Scrivila come 15/10/2026.")


def _occupazioni_giorno(cli, giorno, sede):
    """Tutte le righe aula/fascia di un giorno: (righe occupate, {aula: edificio})."""
    cli.get(URL_SPAZI, {"evn_init": "event", "jaf_currentWFID": "main"})  # apre la sessione
    page = soup(cli.get(URL_SPAZI, {
        "csic": sede, "categoria": "tutte", "tipologia": "tutte", "giorno_day": giorno.day,
        "giorno_month": giorno.month, "giorno_year": giorno.year, "jaf_giorno_date_format": "dd/MM/yyyy",
        "evn_visualizza": "", "jaf_currentWFID": "main"}))
    righe, aule, edificio, inizio_griglia = [], {}, None, 8 * 60
    for tr in page.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if not tds:
            continue
        if "innerEdificio" in (tds[0].get("class") or []):
            edificio = txt(tds[0])
            continue
        if any("innerOrario" in (td.get("class") or []) for td in tds):
            # le ore sono centrate sul loro istante: «09:00» copre da 08:30 a 09:30
            col = 0
            for td in tds:
                span = int(td.get("colspan", 1) or 1)
                if "innerOrario" in (td.get("class") or []):
                    if re.fullmatch(r"\d{2}:\d{2}", txt(td)):
                        inizio_griglia = _minuti(txt(td)) - (col + span / 2) * 15
                        break
                    col += span
            inizio_griglia = int(inizio_griglia)
            continue
        if len(tds) < 3 or "data" not in (tds[0].get("class") or []) or "dove" not in (tds[1].get("class") or []):
            continue
        a = tds[1].find("a")
        aula = txt(tds[1])
        aule[aula] = (a.get("title") if a else None) or edificio
        t = inizio_griglia
        for td in tds[2:]:
            span = int(td.get("colspan", 1) or 1)
            if "slot" in (td.get("class") or []):
                descr = re.sub(r"\s+", " ", txt(td))
                m = re.match(r"^(.*?)\s+(\d{6})\s*-?\s*(.*)$", descr)
                righe.append({
                    "data": giorno.strftime("%d/%m/%Y"), "giorno": sm.GIORNI[giorno.weekday()],
                    "aula": aula, "inizio": _hhmm(t), "fine": _hhmm(t + 15 * span),
                    "insegnamento": m.group(1) if m else descr, "codice": m.group(2) if m else None,
                    "docente": _nome_proprio(m.group(3)) if m else None,
                    "edificio": aule[aula], "descrizione": descr})
            t += 15 * span
    return righe, aule


def _intervalli_liberi(occupati, da, a):
    """Gli intervalli liberi tra `da` e `a` (minuti), dati gli intervalli occupati."""
    liberi, t = [], da
    for x, y in sorted(occupati):
        if x > t:
            liberi.append((t, min(x, a)))
        t = max(t, y)
        if t >= a:
            break
    if t < a:
        liberi.append((t, a))
    return [(x, y) for x, y in liberi if y > x]


_aule = {}  # sede -> nomi delle aule


def aule_sede(sede, log=None, stop=None):
    """I nomi delle aule di una sede del sito Spazi (dalla griglia di oggi), in ordine naturale."""
    if sede not in _aule:
        _, aule = _occupazioni_giorno(_cli(log, stop), date.today(), sede)
        _aule[sede] = sorted(aule, key=lambda a: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", a)])
    return _aule[sede]


def occupazione_aule(giorno, al=None, sede=None, aula=None, testo=None, dalle=None, alle=None, libere=False,
                     log=None, stop=None, avanzamento=None):
    """Chi occupa le aule di una sede, giorno per giorno (dal sito Spazi). Filtri: `aula` (es. T.2.2),
    `testo` (parole nella descrizione: insegnamento, codice o docente), `dalle`/`alle` (ore).
    Con libere=True elenca invece le aule libere (tra `dalle` e `alle`, altrimenti 08:00–20:00)."""
    if not sede:
        raise ValueError("Scegli la sede (es. MIA = Milano Città Studi, MIB = Milano Bovisa).")
    g1 = _data(giorno) if isinstance(giorno, str) else giorno
    g2 = (_data(al) if isinstance(al, str) else al) if al else g1
    if g2 < g1:
        raise ValueError("La data finale viene prima di quella iniziale.")
    if (g2 - g1).days >= MAX_GIORNI_AULE:
        raise ValueError(f"Al massimo {MAX_GIORNI_AULE} giorni per volta.")
    da_min = _minuti(dalle) if dalle else None
    a_min = _minuti(alle) if alle else None
    if da_min is not None and a_min is not None and a_min <= da_min:
        raise ValueError("L'ora finale viene prima di quella iniziale.")
    avanzamento = avanzamento or (lambda fatti, totale: None)
    cli = _cli(log, stop)
    giorni = [g1 + timedelta(days=k) for k in range((g2 - g1).days + 1)]
    occupate, liberi = [], []
    for n, g in enumerate(giorni, 1):
        righe, aule = _occupazioni_giorno(cli, g, sede)
        avanzamento(n, len(giorni))
        if aula:
            righe = [r for r in righe if r["aula"].lower() == aula.strip().lower()]
            aule = {k: v for k, v in aule.items() if k.lower() == aula.strip().lower()}
        if libere:
            x, y = da_min if da_min is not None else 8 * 60, a_min if a_min is not None else 20 * 60
            for nome, edificio in aule.items():
                occ = [(_minuti(r["inizio"]), _minuti(r["fine"])) for r in righe if r["aula"] == nome]
                lib = _intervalli_liberi(occ, x, y)
                if dalle or alle:  # libera per tutto l'intervallo chiesto (a meno della tolleranza)
                    lib = [(p, q) for p, q in lib if p <= x + TOLLERANZA and q >= y - TOLLERANZA]
                if lib:
                    liberi.append({"data": g.strftime("%d/%m/%Y"), "giorno": sm.GIORNI[g.weekday()], "aula": nome,
                                   "libera": " · ".join(f"{_hhmm(p)}–{_hhmm(q)}" for p, q in lib),
                                   "edificio": edificio})
            continue
        for r in righe:
            if testo and not corrisponde(testo, r["descrizione"]):
                continue
            if da_min is not None and _minuti(r["fine"]) <= da_min + TOLLERANZA:
                continue  # finisce prima (o appena dopo) l'inizio della fascia
            if a_min is not None and _minuti(r["inizio"]) >= a_min - TOLLERANZA:
                continue
            occupate.append(r)
    nome_sede = dict(scelte()["sedi_aule"]).get(sede, sede)
    quando = g1.strftime("%d/%m/%Y") + (f" – {g2.strftime('%d/%m/%Y')}" if g2 != g1 else "")
    if libere:
        fascia = f" dalle {dalle or '08:00'} alle {alle or '20:00'}"
        ris = Risultato(f"Aule libere – {nome_sede}, {quando}{fascia}",
                        [Tabella("Aule libere", ["data", "giorno", "aula", "libera", "edificio"], liberi)])
        ris.note.append("Libere secondo il sito Spazi: un'aula può comunque essere chiusa o riservata.")
        if not liberi:
            ris.note.append("Nessuna aula libera con queste condizioni.")
        return ris
    ris = Risultato(f"Occupazione aule – {nome_sede}, {quando}", [Tabella("Occupazione aule", [
        "data", "giorno", "aula", "inizio", "fine", "insegnamento", "codice", "docente", "edificio",
        "descrizione"], occupate)])
    if not occupate:
        ris.note.append("Nessuna occupazione con queste condizioni (la domenica e i giorni festivi le aule sono vuote).")
    return ris


# ============================================================ informazioni su un corso

_corsi = {}  # (anno, scuola) -> (scuole, corsi)


def corsi_di_studio(aa=None, scuola=None, log=None, stop=None):
    """Le scuole [(codice, nome)] e i corsi di studio della `scuola` [(codice, nome, tipo di laurea)],
    dai menu delle pagine informative. Senza `scuola`: quella proposta dal sito, che è la prima."""
    aa = _aa(aa)
    if (aa, scuola) not in _corsi:
        richiesta = {"evn_default": "EVENTO", "aa": aa, "lang": "IT", **({"k_cf": scuola} if scuola else {})}
        page = soup(_cli(log, stop).get(PAGINE_CORSO["struttura"][1], richiesta))
        scuole = [(o["valore"], o["testo"]) for o in sm.select_options(page, "k_cf")]
        corsi = [(o["valore"], o["testo"], o["gruppo"]) for o in sm.select_options(page, "k_corso_la")]
        _corsi[(aa, scuola)] = (scuole, corsi)
    return _corsi[(aa, scuola)]


def info_corso(corso, pagina="struttura", aa=None, log=None, stop=None):
    """Una pagina informativa di un corso di studi (codice, es. 531): struttura, docenti,
    interdisciplinari, scambi."""
    if pagina not in PAGINE_CORSO:
        raise ValueError(f"Pagina sconosciuta: {pagina}. Valori: {', '.join(PAGINE_CORSO)}")
    if not str(corso or "").strip().isdigit():
        raise ValueError("Scrivi il codice numerico del corso di studi (es. 531 per Ingegneria Informatica).")
    aa = _aa(aa)
    nome, url = PAGINE_CORSO[pagina]
    page = soup(_cli(log, stop).get(url, {"evn_default": "EVENTO", "aa": aa, "k_corso_la": str(corso).strip(),
                                          "lang": "IT"}))
    sel = next((o["testo"] for o in sm.select_options(page, "k_corso_la") if o["selezionata"]), None)
    ris = Risultato(f"{nome} – {sel or 'corso ' + str(corso)} – {aa}/{int(aa) + 1}")
    for titolo, _, righe in _blocchi_pagina(page):
        if pagina == "docenti":
            for r in righe:  # il codice del docente, per aprirne la scheda
                cella = r["_celle"].get("Docente")
                a = cella.find("a", href=True) if cella else None
                r["codice_docente"] = _k_doc(a["href"]) if a else None
        if pagina == "scambi":  # i blocchi sono i Paesi
            righe = [{"paese": r.get("gruppo"), **{k: v for k, v in r.items() if k != "gruppo"}} for r in righe]
        colonne, pulite, unico = _pulisci(righe)
        ris.tabelle.append(Tabella(unico or (nome if len(ris.tabelle) == 0 else titolo) or nome, colonne, pulite))
    ris.note += _messaggi_modulo(page)
    if not ris.tabelle and not ris.note:
        ris.note.append("La pagina non contiene dati per questo corso.")
    return ris


def vecchi_ordinamenti(insegnamento="", docente="", log=None, stop=None):
    """Insegnamenti degli ordinamenti precedenti al D.M. 509 (lauree «vecchio ordinamento»)."""
    if len((insegnamento or "").strip()) < 3 and len((docente or "").strip()) < 3:
        raise ValueError("Scrivi almeno 3 lettere dell'insegnamento o del docente.")
    cli = _cli(log, stop)
    cli.get(URL_VO, {"evn_default": "EVENTO", "lang": "IT", "jaf_currentWFID": "main"})
    page = soup(cli.post(URL_VO + "?jaf_currentWFID=main", {"codDescr": insegnamento or "", "n_docente": docente or "",
                                                            "lang": "IT", "evn_default": "Esegui Ricerca"}))
    ris = Risultato("Ordinamenti precedenti al D.M. 509 – " + ", ".join(
        x for x in (insegnamento and f"insegnamento «{insegnamento}»", docente and f"docente «{docente}»") if x))
    for t in page.find_all("table", class_="TableDati"):
        colonne, righe, _ = _pulisci(righe_tabella(t))
        if righe:
            ris.tabelle.append(Tabella("Insegnamenti", colonne, righe))
    if not ris.tabelle:
        ris.note.append("Nessun insegnamento trovato.")
    return ris


# ============================================================ uscita

def salva(ris, path, formato=None):
    """Salva il risultato. Excel: un foglio per tabella; gli altri formati: la tabella principale."""
    formato = formato or str(path).rsplit(".", 1)[-1].lower()
    if formato not in ex.FORMATI:
        raise ValueError(f"Formato sconosciuto: {formato}. Valori: {', '.join(ex.FORMATI)}")
    tabelle = [t for t in ris.tabelle if t.righe] or ris.tabelle[:1]
    if not tabelle:
        raise ValueError("Non ci sono risultati da salvare.")
    if formato == "xlsx":
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        fogli, usati = {}, set()
        for t in tabelle:
            nome = re.sub(r"[\[\]:*?/\\]", " ", t.nome)[:31].strip() or "Risultati"
            k = 2
            while nome in usati:
                nome = f"{nome[:28]} {k}"
                k += 1
            usati.add(nome)
            fogli[nome] = (t.colonne, t.righe)
        ex.esporta_xlsx(path, fogli)
    else:
        t = tabelle[0]
        ex.esporta(path, formato, t.nome, t.colonne, t.righe, ris.titolo)
    return path


def _valore(v):
    return "" if v is None else str(v).replace("\n", " / ")


def stampa(ris, file=None):
    """Il risultato come testo leggibile nel terminale."""
    file = file or sys.stdout
    larghezza = shutil.get_terminal_size((120, 20)).columns
    print(f"\n{ris.titolo}\n{'=' * min(len(ris.titolo), larghezza)}", file=file)
    for n in ris.note:
        print(f"• {n}", file=file)
    for t in ris.tabelle:
        print(f"\n── {t.nome} ({len(t.righe)} righe)", file=file)
        if not t.righe:
            continue
        cols = [c for c in t.colonne if not c.startswith("url")] or t.colonne
        vals = [[_valore(r.get(c)) for c in cols] for r in t.righe]
        w = [min(max([len(ex.label(c))] + [len(v[i]) for v in vals]), 45) for i, c in enumerate(cols)]
        if sum(w) + 2 * len(w) <= larghezza:
            taglia = lambda s, n: s if len(s) <= n else s[:n - 1] + "…"  # noqa: E731
            print("  ".join(taglia(ex.label(c), w[i]).ljust(w[i]) for i, c in enumerate(cols)), file=file)
            print("  ".join("─" * x for x in w), file=file)
            for v in vals:
                print("  ".join(taglia(s, w[i]).ljust(w[i]) for i, s in enumerate(v)).rstrip(), file=file)
        else:  # troppe colonne per lo schermo: una scheda per riga
            ln = max(len(ex.label(c)) for c in cols)
            for v in vals:
                for c, s in zip(cols, v):
                    if s:
                        print(f"  {ex.label(c).ljust(ln)}  {s}", file=file)
                print(file=file)


# ============================================================ terminale

def main(argv=None):
    for flusso in (sys.stdout, sys.stderr):  # la console di Windows non sempre sa scrivere «–» o «→»
        try:
            flusso.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(
        description="Ricerche sul sito del Politecnico di Milano: chi insegna cosa, quando e dove.",
        epilog="Esempi:\n" + "\n".join(l.strip() for l in __doc__.split("Uso da terminale, esempi:")[1]
                                       .split("Ogni comando")[0].strip().splitlines()),
        formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="comando", metavar="COMANDO", required=True)

    def comando(nome, aiuto, anno=True):
        p = sub.add_parser(nome, help=aiuto, description=aiuto)
        if anno:
            p.add_argument("--aa", help="anno accademico, es. 2026 (= 2026/2027); se omesso quello attuale")
        p.add_argument("--out", help="salva i risultati: .xlsx (tutte le tabelle), .csv, .html o .json")
        return p

    p = comando("chi-insegna", "i docenti di un insegnamento con il loro orario, e chi copre certe fasce orarie")
    p.add_argument("insegnamento", help="nome o codice, es. \"geometria e algebra lineare\" o 082747")
    p.add_argument("--fasce", default="", help="es. \"gio 08:15-10:15, ven 10:15-13:15\" (anche \"gio 8-10\": "
                   f"c'è un margine di {TOLLERANZA} minuti): un intervallo va coperto tutto da una lezione; "
                   "un'ora sola (\"gio 08:15\") vuol dire «a lezione in quel momento», il giorno da solo "
                   "(\"mar\") «a lezione quel giorno»")
    p.add_argument("--sede", help="sede dei manifesti: MI, BV, CO, CR, LC, MN, PC")

    p = comando("insegnamenti", "insegnamenti con i loro docenti")
    p.add_argument("testo", nargs="?", default="", help="parte del nome o del codice dell'insegnamento")
    p.add_argument("--docente", default="", help="parte del nome del docente")
    p.add_argument("--sede", help="sede: MI, BV, CO, CR, LC, MN, PC")

    p = comando("docente", "scheda di un docente: insegnamenti, scaglioni, orario (o elenco se il nome è ambiguo)")
    p.add_argument("chi", help="codice del docente (es. 123456) oppure parte del nome")

    p = comando("aule", "occupazione delle aule giorno per giorno (sito Spazi), o le aule libere", anno=False)
    p.add_argument("--sede", help="es. MIA (Milano Città Studi), MIB (Bovisa); elenco con --elenca-sedi")
    p.add_argument("--giorno", help="gg/mm/aaaa")
    p.add_argument("--al", help="ultimo giorno, gg/mm/aaaa (al massimo 14 giorni)")
    p.add_argument("--aula", help="solo questa aula, es. T.2.2")
    p.add_argument("--cerca", help="parole nella descrizione: insegnamento, codice o docente")
    p.add_argument("--dalle", help="ora, es. 10:15")
    p.add_argument("--alle", help="ora, es. 12:15")
    p.add_argument("--libere", action="store_true", help="elenca le aule libere (tra --dalle e --alle)")
    p.add_argument("--elenca-sedi", action="store_true", help="mostra i codici delle sedi ed esce")

    p = comando("corso", "informazioni su un corso di studi")
    p.add_argument("corso", nargs="?", help="codice del corso, es. 531 (i codici con --elenca)")
    p.add_argument("--elenca", action="store_true", help="elenca le scuole e i corsi di studio con il loro codice")
    p.add_argument("--scuola", help="con --elenca: codice della scuola di cui elencare i corsi")
    p.add_argument("--mostra", choices=list(PAGINE_CORSO), default="struttura",
                   help="struttura (piani di studio), docenti, interdisciplinari, scambi")

    p = comando("vecchi-ordinamenti", "insegnamenti degli ordinamenti precedenti al D.M. 509", anno=False)
    p.add_argument("--insegnamento", default="")
    p.add_argument("--docente", default="")

    a = ap.parse_args(argv)
    log = lambda m: print(m, file=sys.stderr)  # noqa: E731
    try:
        if a.comando == "chi-insegna":
            ris = chi_insegna(a.insegnamento, a.fasce, a.aa, a.sede, log, avanzamento=lambda n, t: print(
                f"  letti {n} docenti su {t}", end="\n" if n == t else "\r", file=sys.stderr))
        elif a.comando == "insegnamenti":
            ris = insegnamenti(a.testo, a.docente, a.aa, a.sede)
        elif a.comando == "docente":
            ris = scheda_docente(a.chi, a.aa)
        elif a.comando == "aule":
            if a.elenca_sedi or not a.sede or not a.giorno:
                if not a.elenca_sedi:
                    print("Servono --sede e --giorno. Sedi disponibili:", file=sys.stderr)
                for codice, nome in scelte()["sedi_aule"]:
                    print(f"  {codice:<6} {nome}")
                return 0 if a.elenca_sedi else 2
            ris = occupazione_aule(a.giorno, a.al, a.sede, a.aula, a.cerca, a.dalle, a.alle, a.libere)
        elif a.comando == "corso":
            if a.elenca or not a.corso:
                scuole, corsi = corsi_di_studio(a.aa, a.scuola)
                print("Scuole (per --scuola):")
                for codice, nome in scuole:
                    print(f"  {codice:<6} {nome}")
                print("\nCorsi di studio:")
                for codice, nome, tipo in corsi:
                    print(f"  {codice:<6} {nome} – {tipo}")
                return 0 if a.elenca else 2
            ris = info_corso(a.corso, a.mostra, a.aa)
        else:
            ris = vecchi_ordinamenti(a.insegnamento, a.docente)
    except ValueError as e:
        print(f"Errore: {e}", file=sys.stderr)
        return 2
    except ErroreSito as e:
        print(f"Il sito non risponde: {e}. Riprova più tardi.", file=sys.stderr)
        return 3
    except Interrotto:
        return 130
    except Exception as e:  # rete assente, sito irraggiungibile…
        print(f"Ricerca non riuscita: {e}", file=sys.stderr)
        return 1
    stampa(ris)
    if a.out:
        salva(ris, a.out)
        print(f"\nSalvato in {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
