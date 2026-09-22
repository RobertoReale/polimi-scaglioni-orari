#!/usr/bin/env python3
"""
Motore dello scraper dei Manifesti degli Studi del Politecnico di Milano
https://onlineservices.polimi.it/manifesti/manifesti/controller/ManifestoPublic.do

Per ogni combinazione  scuola -> corso di studi -> anno -> piano di studio  scarica
l'elenco degli insegnamenti e, per ciascuno, gli scaglioni (docenti/moduli)
e l'orario didattico. Salva tutto in un unico file JSON.

Per l'uso normale avvia l'interfaccia grafica (avvia.bat / avvia.sh / python avvia.py).
Uso da terminale, esempi:
    python scarica_manifesti.py --elenca                      (mostra corsi e codici disponibili)
    python scarica_manifesti.py --sede MI --ordinamento 96/23 --anni-corso 1 --periodi annuale,1sem
    python scarica_manifesti.py --corsi 531 --piani primo --no-orari
"""
import argparse
import hashlib
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

HOST = "https://onlineservices.polimi.it"
BASE = HOST + "/manifesti/manifesti/controller/ManifestoPublic.do"
HERE = Path(__file__).resolve().parent


# --------------------------------------------------------------------------- HTTP

class Client:
    """Sessione HTTP con pausa tra le richieste, retry e cache su disco.
    Si può usare da più thread insieme (una sessione per thread)."""

    def __init__(self, delay=0.4, cache_dir=None, retries=4, log=None, stop=None):
        self._locale = threading.local()
        self._lock = threading.Lock()
        self.delay = delay
        self.retries = retries
        self.cache_dir = Path(cache_dir) if cache_dir else None
        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.log = log or print_log
        self.stop = stop  # threading.Event opzionale: se impostato, interrompe
        self.n_requests = 0

    @property
    def s(self):
        s = getattr(self._locale, "s", None)
        if s is None:
            s = self._locale.s = requests.Session()
            s.headers["User-Agent"] = "Mozilla/5.0 (polimi-scaglioni-orari)"
        return s

    def _pausa(self, secondi):
        """Attende `secondi`, ma si sveglia subito se l'utente interrompe."""
        if self.stop is None:
            time.sleep(secondi)
        elif self.stop.wait(secondi):
            raise Interrotto()

    def get(self, url, params=None):
        req = requests.Request("GET", url, params=params).prepare()
        key = hashlib.sha1(req.url.encode()).hexdigest()
        f = self.cache_dir / f"{key}.html" if self.cache_dir else None
        if f and f.exists():
            return f.read_text(encoding="utf-8")
        for attempt in range(1, self.retries + 1):
            if self.stop is not None and self.stop.is_set():
                raise Interrotto()
            try:
                self._pausa(self.delay)
                r = self.s.get(req.url, timeout=60)
                with self._lock:
                    self.n_requests += 1
                r.raise_for_status()
                r.encoding = r.encoding or "utf-8"
                html = r.text
                break
            except requests.RequestException as e:
                if attempt == self.retries:
                    raise
                wait = 2 ** attempt
                self.log(f"    ! errore di rete ({e}); riprovo tra {wait}s")
                self._pausa(wait)
        if f:
            # scrittura atomica: un'interruzione non lascia file a metà nella cache
            tmp = f.with_suffix(f".{threading.get_ident()}.tmp")
            tmp.write_text(html, encoding="utf-8")
            tmp.replace(f)
        return html


class Interrotto(Exception):
    """Scaricamento fermato dall'utente."""


def print_log(msg):
    print(msg, flush=True)


def soup(html):
    return BeautifulSoup(html, "html.parser")


def txt(el):
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip() if el else ""


def link_pulito(href):
    """URL assoluto senza ';jsessionid=…' (legato alla sessione, scade)."""
    return re.sub(r";jsessionid=[^?#]*", "", urljoin(HOST, href.strip()))


def qs_of(href):
    return {k: v[0] for k, v in parse_qs(urlparse(href).query).items()}


# --------------------------------------------------------------------- tabelle

def own_rows(table):
    """Le <tr> della tabella, escluse quelle di tabelle annidate."""
    return [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]


def table_grid(table):
    """Ricostruisce la griglia di una tabella HTML gestendo rowspan/colspan.
    Ritorna una lista di righe; ogni riga è una lista di tuple
    (colonna_iniziale, colspan, cella, ereditata_da_rowspan)."""
    grid, pending = [], {}  # pending: col -> [cella, colspan, righe_rimanenti]
    for tr in own_rows(table):
        row, col = [], 0
        cells = list(tr.find_all(["td", "th"], recursive=False))
        while cells or any(c >= col for c in pending):
            if col in pending:
                cell, span, left = pending[col]
                row.append((col, span, cell, True))
                if left <= 1:
                    del pending[col]
                else:
                    pending[col][2] -= 1
                col += span
                continue
            if not cells:
                col += 1
                continue
            cell = cells.pop(0)
            span = int(cell.get("colspan", 1) or 1)
            rs = int(cell.get("rowspan", 1) or 1)
            row.append((col, span, cell, False))
            if rs > 1:
                pending[col] = [cell, span, rs - 1]
            col += span
        grid.append(row)
    return grid


def header_names(grid, n_header_rows):
    """Nome di ogni colonna combinando le righe di intestazione."""
    names = {}
    for row in grid[:n_header_rows]:
        for col, span, cell, _ in row:
            t = txt(cell)
            for c in range(col, col + span):
                if t and t not in names.get(c, ""):
                    names[c] = (names.get(c, "") + " " + t).strip()
    return names


def cell_info(cell):
    """Testo della cella + eventuali link (testo, url) e immagini-bandiera."""
    info = {"testo": txt(cell)}
    links = [(txt(a), link_pulito(a["href"])) for a in cell.find_all("a", href=True) if txt(a)]
    if links:
        info["link"] = links
    flags = [Path(i["src"]).stem for i in cell.find_all("img", src=True) if "flag" in i["src"]]
    if flags:
        info["lingue"] = flags
    return info


# ------------------------------------------------------------ pagina manifesto

def manifesto_page(cli, **params):
    p = {"evn_default": "EVENTO", "lang": "IT"}
    p.update({k: v for k, v in params.items() if v is not None})
    return soup(cli.get(BASE, p))


def select_options(page, name):
    """Opzioni di una <select>: lista di dict {valore, testo, gruppo, selezionata}."""
    sel = page.find("select", attrs={"name": name})
    out = []
    if not sel:
        return out
    for o in sel.find_all("option"):
        g = o.find_parent("optgroup")
        out.append({
            "valore": o.get("value", ""),
            "testo": txt(o),
            "gruppo": g.get("label", "") if g else "",
            "selezionata": o.has_attr("selected"),
        })
    return out


def piano_info(page):
    """Testo 'Sede: ... / Lingua Offerta: ...' mostrato sotto la select dei piani."""
    sel = page.find("select", attrs={"name": "k_indir"})
    td = sel.find_parent("td") if sel else None
    t = td.get_text("\n", strip=True) if td else ""
    sede = re.search(r"Sede:\s*(.+)", t)
    lingua = re.search(r"Lingua Offerta:\s*(.+)", t)
    return (sede.group(1).strip() if sede else None,
            lingua.group(1).strip() if lingua else None)


MANIFESTO_KEYS = {
    "Codice": "codice",
    "SSD SM 639/24": "ssd_639_24",
    "SSD": "ssd",
    "Denominazione Insegnamento": "nome",
    "Num Sez": "num_sez",
    "Lingua": "lingua",
    "Sede d'erogazione": "sede_erogazione",
    "Tipo": "tipo",
    "Periodo": "periodo",
    "CFU": "cfu",
    "CFU Gruppo": "cfu_gruppo",
}


def parse_insegnamenti(page):
    """Tutte le righe del manifesto che hanno un link al dettaglio."""
    out, seen = [], set()
    for table in page.find_all("table", class_="TableDati"):
        grid = table_grid(table)
        if not grid or not any(txt(c) == "Codice" for _, _, c, _ in grid[0]):
            continue
        names = header_names(grid, 1)
        # titolo del blocco (es. "1°Anno") = TitleInfoCard che precede la tabella
        title = table.find_previous("td", class_="TitleInfoCard")
        for row in grid[1:]:
            link = None
            for _, _, cell, _ in row:
                a = cell.find("a", href=re.compile("EVN_DETTAGLIO_RIGA_MANIFESTO"))
                if a and txt(a):
                    link = a
                    break
            if not link:
                continue
            url = link_pulito(link["href"])
            if url in seen:
                continue
            seen.add(url)
            q = qs_of(url)
            ins = {}
            for col, _, cell, _ in row:
                h = names.get(col, f"col{col}")
                if "Orario" in h:
                    continue
                key = MANIFESTO_KEYS.get(h, h)
                info = cell_info(cell)
                if key == "lingua":
                    ins[key] = info.get("lingue", [info["testo"]] if info["testo"] else [])
                elif key == "cfu":
                    m = re.match(r"([\d.,]+)\s*(?:\[([\d.,]+))?", info["testo"])
                    ins["cfu"] = m.group(1) if m else info["testo"]
                    if m and m.group(2):
                        ins["cfu_didattica_innovativa"] = m.group(2)
                else:
                    ins[key] = info["testo"]
            ins["nome"] = txt(link)
            ins["blocco"] = txt(title) if title else None
            ins["anno_corso"] = q.get("anno_corso")
            ins["semestre_codice"] = q.get("semestre")
            ins["id_item_offerta"] = q.get("idItemOfferta")
            ins["id_riga"] = q.get("idRiga")
            ins["url_dettaglio"] = url
            out.append(ins)
    return out


def norm_periodo(p):
    p = (p or "").lower().replace("°", "").replace(" ", "")
    return {"1": "1sem", "2": "2sem", "a": "annuale", "ann": "annuale"}.get(p, p)


# ---------------------------------------------------------- dettaglio / scaglioni

SCAGLIONI_KEYS = {
    "Scaglione Da (compreso)": "da",
    "Scaglione A (escluso)": "a",
    "Codice": "codice_modulo",
    "Denominazione Modulo": "modulo",
    "Docente/i": "docenti",
    "CFU": "cfu",
    "Periodo": "periodo",
    "Lingua offerta": "lingua",
    "Programma dettagliato": "programma",
}


def parse_scaglioni(fragment):
    """Tabella scaglioni del tab 'Dettaglio' -> lista di scaglioni con righe."""
    table = None
    for t in fragment.find_all("table"):
        rows = own_rows(t)
        if rows and "Scaglione" in txt(rows[0]):
            table = t
            break
    if table is None:
        return [], []
    grid = table_grid(table)
    n_head = 2 if len(grid) > 1 and any("Da" in txt(c) for _, _, c, _ in grid[1]) else 1
    names = header_names(grid, n_head)
    righe = []
    for row in grid[n_head:]:
        r = {}
        for col, _, cell, _ in row:
            h = names.get(col, f"col{col}")
            if "Orario" in h:
                continue
            key = SCAGLIONI_KEYS.get(h, h)
            info = cell_info(cell)
            if key == "docenti":
                r["docenti"] = [l[0] for l in info.get("link", [])] or ([info["testo"]] if info["testo"] else [])
                r["docenti_url"] = [l[1] for l in info.get("link", [])]
            elif key == "lingua":
                r["lingua"] = info.get("lingue", [])
            elif key == "programma":
                r["programma_url"] = [l[1] for l in info.get("link", [])] or [
                    link_pulito(a["href"]) for a in cell.find_all("a", href=True)]
            else:
                r[key] = info["testo"]
        if any(v for v in r.values()):
            righe.append(r)
    # raggruppa per scaglione (da, a)
    scaglioni, idx = [], {}
    for r in righe:
        k = (r.get("da"), r.get("a"))
        if k not in idx:
            idx[k] = len(scaglioni)
            scaglioni.append({"da": k[0], "a": k[1], "docenti": [], "righe": []})
        sc = scaglioni[idx[k]]
        sc["righe"].append({kk: vv for kk, vv in r.items() if kk not in ("da", "a")})
        for d in r.get("docenti", []):
            if d not in sc["docenti"]:
                sc["docenti"].append(d)
    return scaglioni, righe


# --------------------------------------------------------------------- orario

GIORNI = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]


MESI = {"gen": 1, "feb": 2, "mar": 3, "apr": 4, "mag": 5, "giu": 6,
        "lug": 7, "ago": 8, "set": 9, "ott": 10, "nov": 11, "dic": 12}


def parse_data_breve(s):
    """'14-set-26' -> '14/09/2026' (None se non riconosciuta)."""
    m = re.match(r"(\d{1,2})-([a-z]{3})-(\d{2})$", s.strip().lower())
    if not m or m.group(2) not in MESI:
        return None
    return f"{int(m.group(1)):02d}/{MESI[m.group(2)]:02d}/20{m.group(3)}"


def hhmm(minutes):
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def parse_orario_grid(table):
    """Una griglia oraria (table.scrollTable) -> lista di lezioni."""
    rows = own_rows(table)
    if not rows:
        return []
    # calibrazione asse orario dalla riga di intestazione
    head = rows[0].find_all("td", recursive=False)
    col, base_min, min_per_col = 0, None, 15
    for td in head:
        span = int(td.get("colspan", 1) or 1)
        if "innerDataDove" in (td.get("class") or []):
            continue
        m = re.match(r"(\d{1,2}):(\d{2})", txt(td))
        if m and base_min is None:
            min_per_col = 60 // span
            # l'etichetta è centrata sulla linea dell'ora
            base_min = int(m.group(1)) * 60 + int(m.group(2)) - (col + span // 2) * min_per_col
        col += span
    if base_min is None:
        base_min = 8 * 60
    lezioni, giorno, aula, aula_desc, aula_url = [], None, None, None, None
    for tr in rows[1:]:
        if "riga_vuota" in (tr.get("class") or []):
            continue
        col = 0
        for td in tr.find_all("td", recursive=False):
            cls = td.get("class") or []
            if "data" in cls:
                giorno = txt(td)
                continue
            if "dove" in cls:
                a = td.find("a")
                aula = txt(td) or None
                aula_desc = a.get("title") if a else None
                aula_url = a.get("href", "").strip() if a else None
                continue
            span = int(td.get("colspan", 1) or 1)
            if "slot" in cls:
                a = td.find("a")
                title = ((a.get("title") if a else None) or txt(td)).strip()
                date = re.findall(r"\d{2}/\d{2}/\d{4}", title)
                # formato alternativo: elenco puntato delle singole date (es. 14-set-26)
                date_lez = [d for d in (parse_data_breve(txt(li)) for li in td.find_all("li")) if d]
                if date_lez and not date:
                    date = [date_lez[0], date_lez[-1]]
                sem = next((c.replace("slotSem", "") for c in cls if c.startswith("slotSem")), None)
                start = base_min + col * min_per_col
                lezioni.append({
                    "giorno": giorno,
                    "giorno_n": GIORNI.index(giorno) + 1 if giorno in GIORNI else None,
                    "inizio": hhmm(start),
                    "fine": hhmm(start + span * min_per_col),
                    "durata_min": span * min_per_col,
                    "aula": aula,
                    "aula_descrizione": aula_desc,
                    "aula_url": aula_url,
                    "attivita": re.sub(r"\s*\(dal.*$", "", title).strip(),
                    "dal": date[0] if date else None,
                    "al": date[1] if len(date) > 1 else None,
                    "date_lezioni": date_lez or None,
                    "semestre": sem or None,
                })
            col += span
    return lezioni


def parse_orario(fragment):
    """Frammento del tab 'Orario didattico' -> lista di lezioni con scaglione."""
    lezioni, sc_da, sc_a, periodo = [], None, None, None
    for t in fragment.find_all("table"):
        cls = t.get("class") or []
        if "SimpleCardOut" in cls and not t.find("table"):
            intest = txt(t)
            m = re.search(r"da \(compreso\):\s*(\S+).*?a \(escluso\):\s*(\S+)", intest)
            if m:
                sc_da, sc_a = m.group(1), m.group(2)
            p = re.search(r"((?:Primo|Secondo|Terzo) Semestre|Insegnamento Annuale|Annuale)", intest)
            if p:
                periodo = p.group(1)
        elif "scrollTable" in cls:
            for lz in parse_orario_grid(t):
                lz.update({"scaglione_da": sc_da, "scaglione_a": sc_a, "periodo_orario": periodo})
                lezioni.append(lz)
    return lezioni


NOTA_ORARIO_ASSENTE = "orario non pubblicato sul sito"
NOTA_NESSUNA_LEZIONE = "nessuna lezione in orario"


def nota_orario_vuoto(testo):
    """Messaggio leggibile quando il tab orario non contiene lezioni."""
    if not testo or "Non esistono occupazioni" in testo or "Data Dove" in testo:
        return NOTA_NESSUNA_LEZIONE  # messaggio del sito o griglia settimanale vuota
    return testo[:300]


def fetch_dettaglio(cli, url, with_orari):
    """Pagina di dettaglio di un insegnamento: una o più sezioni (tab)."""
    page = soup(cli.get(url))
    sezioni = []
    for tabs in page.find_all("div", class_="tabs"):
        sid = tabs.get("id")
        li_sc = tabs.find("li", id=f"tab_scaglioni_{sid}")
        li_or = tabs.find("li", id=f"tab_orario_{sid}")
        content = page.find(id=f"content_{sid}")
        if content is None or not content.find("table"):
            if li_sc is None:
                continue
            content = soup(cli.get(BASE + "?evn_DETTAGLIO_SCAGLIONI_AJAX=evento" + li_sc["qs"]))
        scaglioni, righe = parse_scaglioni(content)
        sez = {
            "id_sezione": sid,
            "n_scaglioni": len(scaglioni),
            "scaglioni": scaglioni,
        }
        if with_orari and li_or is not None:
            disabled = "ui-state-disabled" in (li_or.get("class") or [])
            if disabled:
                sez["orario"] = []
                sez["orario_nota"] = NOTA_ORARIO_ASSENTE
            else:
                frag = soup(cli.get(BASE + "?evn_DETTAGLIO_ORARIO_AJAX=evento" + li_or["qs"]))
                sez["orario"] = parse_orario(frag)
                if not sez["orario"]:
                    sez["orario_nota"] = nota_orario_vuoto(txt(frag))
        sezioni.append(sez)
    nota = None
    if not sezioni:
        # es. insegnamenti erogati da atenei partner: nessun tab, solo la scheda
        t = txt(page)
        a = max(t.find("Anno di Corso"), 0)
        b = t.find("manifesti v.")
        nota = t[a:b if b > a else a + 1500].strip()[:1500]
    return sezioni, nota


# ------------------------------------------------------------------ parallelismo

# il tempo è quasi tutto attesa del server (1-2 s a pagina): poche richieste
# contemporanee accorciano molto lo scaricamento senza pesare sul sito
PARALLELI_DEFAULT = 4


def in_parallelo(funzione, elementi, paralleli):
    """Esegue funzione(e) per ogni elemento su `paralleli` thread e restituisce
    (indice, elemento, risultato, errore) man mano che i lavori finiscono.
    Se l'utente interrompe, annulla i lavori non ancora partiti e rilancia Interrotto."""
    with ThreadPoolExecutor(max_workers=max(1, int(paralleli))) as pool:
        futuri = {pool.submit(funzione, e): (k, e) for k, e in enumerate(elementi)}
        try:
            for fut in as_completed(futuri):
                k, e = futuri[fut]
                try:
                    res, err = fut.result(), None
                except Interrotto:
                    raise
                except Exception as x:
                    res, err = None, x
                yield k, e, res, err
        finally:
            for fut in futuri:
                fut.cancel()


# -------------------------------------------------------------------- catalogo

def catalogo(aa=None, sede=None, lang="IT", log=print_log, stop=None):
    """Legge dal sito le scelte disponibili: anni accademici, sedi, scuole e corsi.
    Non usa la cache, così riflette sempre lo stato attuale del sito."""
    cli = Client(delay=0.1, log=log, stop=stop)
    home = soup(cli.get(BASE, {"evn_DEFAULT": "evento", "lang": lang}))
    anni_acc = select_options(home, "aa")
    sedi = select_options(home, "sede")
    aa = aa or next((o["valore"] for o in anni_acc if o["selezionata"]), anni_acc[0]["valore"])
    sede = sede or "ALL_SEDI"
    first = manifesto_page(cli, aa=aa, sede=sede, lang=lang)

    def corsi_scuola(sc):
        pg = manifesto_page(cli, aa=aa, sede=sede, k_cf=sc["valore"], lang=lang)
        return [{"codice": o["valore"], "nome": o["testo"], "gruppo": o["gruppo"]}
                for o in select_options(pg, "k_corso_la") if o["valore"]]

    elenco = select_options(first, "k_cf")
    scuole = [None] * len(elenco)
    for k, sc, corsi, err in in_parallelo(corsi_scuola, elenco, PARALLELI_DEFAULT):
        if err:
            raise err
        scuole[k] = {"codice": sc["valore"], "nome": sc["testo"], "corsi": corsi}
    return {
        "anni_accademici": [{"codice": o["valore"], "nome": o["testo"]} for o in anni_acc],
        "sedi": [{"codice": o["valore"], "nome": o["testo"]} for o in sedi],
        "aa": aa, "sede": sede, "lang": lang,
        "scuole": scuole,
    }


# ------------------------------------------------------------------ scaricamento

PERIODI = {"annuale": "Annuale", "1sem": "1° semestre", "2sem": "2° semestre", "altro": "Altri periodi"}


@dataclass
class Opzioni:
    """Cosa scaricare. None nei filtri = nessun filtro (tutto)."""
    aa: str = None                     # anno accademico, es. "2026" (= 2026/2027); None = quello attuale
    sede: str = "ALL_SEDI"             # MI, BV, CO, CR, LC, MN, PC, ALL_SEDI
    lang: str = "IT"                   # solo IT: il parser riconosce le intestazioni in italiano
    scuole: list = None                # codici scuola, es. ["225"]
    corsi: list = None                 # codici corso di studi, es. ["531", "1030"]
    ordinamento: str = ""              # testo nel tipo di laurea, es. "96/23" o "270"
    tipo_laurea: list = None           # es. ["Primo Livello", "Magistrale"]
    anni_corso: list = field(default_factory=lambda: ["0"])  # "1", "2", ... ; "0" = tutti
    piani: object = "tutti"            # "tutti", "primo" oppure lista di codici es. ["IT1"]
    includi_non_diversificato: bool = False
    includi_altre_sedi: bool = False   # tieni anche i piani di sedi diverse da quella scelta
    periodi: set = None                # {"annuale","1sem","2sem","altro"}; None = tutti
    scaglioni: bool = True             # apri il dettaglio di ogni insegnamento
    orari: bool = True                 # scarica anche l'orario didattico
    out: str = None                    # file JSON; None = nome automatico in output/
    delay: float = 0.4                 # pausa tra le richieste (secondi)
    paralleli: int = PARALLELI_DEFAULT # richieste contemporanee al sito
    cache: str = str(HERE / "cache")   # cartella cache; "" = disattivata


def periodo_ok(periodo, periodi):
    if periodi is None:
        return True
    p = norm_periodo(periodo)
    return p in periodi or ("altro" in periodi and p not in ("annuale", "1sem", "2sem"))


def nome_file_default(aa, sede):
    return HERE / "output" / f"manifesti_{aa}_{sede}_{datetime.now():%Y-%m-%d_%H%M%S}.json"


def riepilogo(result):
    """Conteggi finali del file: mostrati a fine scaricamento e salvati in meta."""
    ins = [i for c in result["corsi_di_studio"] for p in c["piani"] for i in p["insegnamenti"]]
    return {
        "corsi": len(result["corsi_di_studio"]),
        "piani": sum(len(c["piani"]) for c in result["corsi_di_studio"]),
        "insegnamenti": len(ins),
        "scaglioni": sum(i.get("n_scaglioni") or 0 for i in ins),
        "lezioni_settimanali": sum(len(s.get("orario", [])) for i in ins for s in i.get("sezioni", [])),
        "insegnamenti_con_errore": sum(1 for i in ins if i.get("errore")),
        "corsi_con_errore": sum(1 for c in result["corsi_di_studio"] if c.get("errore")),
    }


def scarica(opt, log=print_log, progress=None, stop=None):
    """Esegue lo scraping secondo `opt`. progress(fase, fatti, totale) è opzionale.
    Ritorna il percorso del file JSON salvato (anche se interrotto: salva il parziale)."""
    progress = progress or (lambda *a: None)
    cli = Client(delay=opt.delay, cache_dir=opt.cache or None, log=log, stop=stop)
    lang = opt.lang
    tipi_f = [t.lower() for t in (opt.tipo_laurea or [])]
    anni = [str(a) for a in opt.anni_corso] or ["0"]

    home = soup(Client(delay=0, log=log, stop=stop).get(BASE, {"evn_DEFAULT": "evento", "lang": lang}))
    anni_acc = select_options(home, "aa")
    if not anni_acc:
        raise ValueError("Il sito non ha restituito gli anni accademici: riprova tra poco.")
    aa = opt.aa or next((o["valore"] for o in anni_acc if o["selezionata"]), anni_acc[0]["valore"])
    sedi = {o["valore"]: o["testo"] for o in select_options(home, "sede")}
    if opt.sede not in sedi:
        raise ValueError(f"Sede '{opt.sede}' non valida. Valori: {', '.join(sedi)}")
    sede_nome = re.sub(r"\s*\(\w+\)$", "", sedi[opt.sede])
    log(f"Anno accademico {aa}/{int(aa) + 1} | sede {sedi[opt.sede]} | "
        f"anni di corso: {', '.join('tutti' if a == '0' else a for a in anni)}")

    params = {k: (sorted(v) if isinstance(v, set) else v) for k, v in asdict(opt).items()}
    result = {
        "meta": {
            "fonte": BASE,
            "generato_il": datetime.now().isoformat(timespec="seconds"),
            "completo": False,
            "parametri": {**params, "aa": aa, "sede_nome": sedi[opt.sede]},
            "note": {
                "orario": "inizio/fine calcolati dalla griglia a quarti d'ora del sito",
                "scaglioni": "da (compreso) / a (escluso) = iniziali del cognome dello studente",
            },
        },
        "corsi_di_studio": [],
    }
    out = Path(opt.out) if opt.out else nome_file_default(aa, opt.sede)

    def salva():
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")

    try:
        # ---- fase 1: corsi da elaborare
        first = manifesto_page(cli, aa=aa, sede=opt.sede, lang=lang)
        lavori = []
        for sc in select_options(first, "k_cf"):
            if opt.scuole and sc["valore"] not in opt.scuole:
                continue
            pg = manifesto_page(cli, aa=aa, sede=opt.sede, k_cf=sc["valore"], lang=lang)
            for c in select_options(pg, "k_corso_la"):
                if not c["valore"] or (opt.corsi and c["valore"] not in opt.corsi):
                    continue
                if opt.ordinamento and opt.ordinamento.lower() not in c["gruppo"].lower():
                    continue
                if tipi_f and not any(t in c["gruppo"].lower() for t in tipi_f):
                    continue
                lavori.append((sc, c))
        log(f"{len(lavori)} corsi di studio da elaborare")

        # ---- fase 2: piani ed elenco insegnamenti (più corsi insieme)
        def elabora_corso(lavoro):
            sc, c = lavoro
            corso = {
                "scuola": {"codice": sc["valore"], "nome": sc["testo"]},
                "codice": c["valore"], "nome": c["testo"], "tipo_ordinamento": c["gruppo"],
                "anni_disponibili": [], "piani": [], "piani_scartati": [], "note": [],
            }
            righe_log = []
            for anno in anni:
                pg = manifesto_page(cli, aa=aa, sede=opt.sede, k_cf=sc["valore"],
                                    k_corso_la=c["valore"], ac_ins=anno, lang=lang)
                corso["anni_disponibili"] = [o["valore"] for o in select_options(pg, "ac_ins")]
                if anno not in corso["anni_disponibili"]:
                    corso["note"].append(f"anno di corso {anno} non disponibile")
                    righe_log.append(f"  anno {anno}: non disponibile")
                    continue
                piani = select_options(pg, "k_indir")
                cand = [p for p in piani if p["valore"] != "***" or opt.includi_non_diversificato
                        or all(q["valore"] == "***" for q in piani)]
                if not piani:  # piano unico: il sito non mostra la select dei piani
                    cand = [{"valore": None, "testo": "(piano unico)"}]
                elif isinstance(opt.piani, (list, tuple, set)):
                    cand = [p for p in cand if p["valore"] in opt.piani]
                presi = 0
                for p in cand:
                    if opt.piani == "primo" and presi:
                        break
                    pp = manifesto_page(cli, aa=aa, sede=opt.sede, k_cf=sc["valore"], k_corso_la=c["valore"],
                                        ac_ins=anno, k_indir=p["valore"], lang=lang,
                                        caricaOffertaComune="on" if opt.includi_non_diversificato else None)
                    p_sede, p_lingua = piano_info(pp)
                    if (not opt.includi_altre_sedi and opt.sede != "ALL_SEDI" and p_sede
                            and sede_nome.lower() not in p_sede.lower()):
                        corso["piani_scartati"].append({"codice": p["valore"], "nome": p["testo"],
                                                        "sede": p_sede, "anno_corso": anno})
                        righe_log.append(f"  anno {anno} · piano {p['valore']} scartato (sede: {p_sede})")
                        continue
                    ins_all = parse_insegnamenti(pp)
                    ins = [i for i in ins_all if periodo_ok(i.get("periodo"), opt.periodi)]
                    corso["piani"].append({
                        "codice": p["valore"], "nome": p["testo"], "sede": p_sede, "lingua": p_lingua,
                        "anno_corso": anno, "n_insegnamenti_totali": len(ins_all), "insegnamenti": ins})
                    presi += 1
                    righe_log.append(f"  anno {'tutti' if anno == '0' else anno} · piano {p['valore'] or 'unico'} "
                                     f"({p_sede or '-'}): {len(ins)} insegnamenti su {len(ins_all)}")
            if not corso["note"]:
                del corso["note"]
            return corso, righe_log

        corsi = result["corsi_di_studio"] = [None] * len(lavori)  # stesso ordine del sito
        progress("Elenco insegnamenti", 0, len(lavori))
        for n, (k, (sc, c), res, err) in enumerate(in_parallelo(elabora_corso, lavori, opt.paralleli), 1):
            log(f"\n[{n}/{len(lavori)}] {c['testo']} ({c['valore']})")
            if err:  # un corso non letto non ferma gli altri
                res = ({"scuola": {"codice": sc["valore"], "nome": sc["testo"]}, "codice": c["valore"],
                        "nome": c["testo"], "tipo_ordinamento": c["gruppo"], "piani": [],
                        "errore": str(err) or repr(err)}, [f"  ! non letto: {err!r}"])
            corsi[k], righe_log = res
            for r in righe_log:
                log(r)
            progress("Elenco insegnamenti", n, len(lavori))
        da_dettagliare = [i for c in corsi for p in c["piani"] for i in p["insegnamenti"]]

        # ---- fase 3: scaglioni e orari (più insegnamenti insieme)
        if opt.scaglioni or opt.orari:
            log(f"\nDettaglio di {len(da_dettagliare)} insegnamenti (scaglioni"
                f"{' e orari' if opt.orari else ''})")
            tot = len(da_dettagliare)
            progress("Scaglioni e orari", 0, tot)
            dettaglio = lambda i: fetch_dettaglio(cli, i["url_dettaglio"], opt.orari)  # noqa: E731
            for n, (_, i, res, err) in enumerate(in_parallelo(dettaglio, da_dettagliare, opt.paralleli), 1):
                progress("Scaglioni e orari", n, tot)
                if err:  # un insegnamento rotto non ferma tutto
                    i["errore"] = str(err) or repr(err)
                    log(f"  ! {i.get('nome')}: {err!r}")
                    continue
                i["sezioni"], nota = res
                if nota:
                    i["nota_dettaglio"] = nota
                i["n_scaglioni"] = sum(s["n_scaglioni"] for s in i["sezioni"])
                n_les = sum(len(s.get("orario", [])) for s in i["sezioni"])
                log(f"  {i.get('codice', '')} {(i.get('nome') or '')[:48]:<48} {i.get('periodo', ''):<8} "
                    f"scaglioni={i['n_scaglioni']:<2} lezioni={n_les}")
        n_err = sum(1 for c in corsi if c.get("errore"))
        if n_err:
            result["meta"]["errore"] = f"{n_err} corsi di studio non letti per un errore (vedi il registro)"
        result["meta"]["completo"] = not n_err
    except Interrotto:
        result["meta"]["interrotto"] = True
        log("\nInterrotto dall'utente: salvo i dati raccolti finora.")
    except Exception as e:  # es. sito irraggiungibile: non perdere quanto già scaricato
        result["meta"]["errore"] = str(e) or repr(e)
        log(f"\nERRORE: {e!r}\nSalvo i dati raccolti finora.")
    # interrotto durante la fase 2: tieni solo i corsi già letti
    result["corsi_di_studio"] = [c for c in result["corsi_di_studio"] if c]
    result["meta"]["riepilogo"] = rie = riepilogo(result)
    salva()
    log(f"\nFatto: {rie['corsi']} corsi, {rie['piani']} piani, {rie['insegnamenti']} insegnamenti, "
        f"{rie['scaglioni']} scaglioni, {rie['lezioni_settimanali']} lezioni settimanali "
        f"({cli.n_requests} pagine scaricate dal sito)."
        + (f"\n{rie['insegnamenti_con_errore']} insegnamenti non letti per errore (vedi colonna Note)."
           if rie["insegnamenti_con_errore"] else "")
        + f"\nSalvato in: {out}")
    return out


# ------------------------------------------------------------------ riga di comando

def csv_arg(v):
    if v is None or v.strip().lower() in ("", "all", "tutti", "tutte"):
        return None
    return [x.strip() for x in v.split(",") if x.strip()]


def main():
    ap = argparse.ArgumentParser(
        description="Scraper dei Manifesti degli Studi PoliMi -> JSON (uso da terminale; "
                    "per l'interfaccia grafica avvia avvia.py)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__)
    ap.add_argument("--elenca", action="store_true",
                    help="mostra anni accademici, sedi, scuole e corsi disponibili ed esce")
    ap.add_argument("--aa", help="anno accademico, es. 2026 (= 2026/2027); se omesso quello attuale")
    ap.add_argument("--sede", default="ALL_SEDI", help="MI, BV, CO, CR, LC, MN, PC o ALL_SEDI (tutte)")
    ap.add_argument("--scuole", default="all", help="codici scuola separati da virgola, o 'all'")
    ap.add_argument("--corsi", default="all", help="codici corso di studi, es. 1030,531, o 'all'")
    ap.add_argument("--ordinamento", default="", help="es. '96/23' o '270'; vuoto = tutti")
    ap.add_argument("--tipo-laurea", default="",
                    help="es. 'Primo Livello', 'Magistrale', 'Ciclo Unico' (virgola per più); vuoto = tutti")
    ap.add_argument("--anni-corso", default="0", help="anni di corso, es. 1 oppure 1,2; 0 = tutti")
    ap.add_argument("--piani", default="tutti", help="'tutti', 'primo' o codici es. IT1,IE1")
    ap.add_argument("--includi-non-diversificato", action="store_true", help="includi il piano '***'")
    ap.add_argument("--includi-altre-sedi", action="store_true",
                    help="tieni anche i piani erogati in sedi diverse da --sede")
    ap.add_argument("--periodi", default="all", help="annuale,1sem,2sem,altro oppure 'all'")
    ap.add_argument("--no-orari", action="store_true", help="non scaricare gli orari")
    ap.add_argument("--no-dettagli", action="store_true", help="solo elenco insegnamenti (niente scaglioni/orari)")
    ap.add_argument("--out", help="file JSON di output (default: output/manifesti_<aa>_<sede>_<data>.json)")
    ap.add_argument("--delay", type=float, default=0.4, help="pausa in secondi tra le richieste")
    ap.add_argument("--paralleli", type=int, default=PARALLELI_DEFAULT,
                    help=f"richieste contemporanee al sito (default {PARALLELI_DEFAULT}; 1 = una alla volta)")
    ap.add_argument("--cache", default=str(HERE / "cache"), help="cartella cache; '' per disattivarla")
    args = ap.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if args.elenca:
        cat = catalogo(args.aa, args.sede)
        print("Anni accademici:", ", ".join(f"{a['codice']} ({a['nome']})" for a in cat["anni_accademici"]))
        print("Sedi:", ", ".join(f"{s['codice']} ({s['nome']})" for s in cat["sedi"]))
        for sc in cat["scuole"]:
            print(f"\nScuola {sc['codice']}: {sc['nome']}")
            for c in sc["corsi"]:
                print(f"  {c['codice']:>6}  {c['nome']}  [{c['gruppo']}]")
        return

    no_det = args.no_dettagli
    periodi = csv_arg(args.periodi)
    opt = Opzioni(
        aa=args.aa, sede=args.sede,
        scuole=csv_arg(args.scuole), corsi=csv_arg(args.corsi),
        ordinamento=args.ordinamento, tipo_laurea=csv_arg(args.tipo_laurea),
        anni_corso=csv_arg(args.anni_corso) or ["0"],
        piani=args.piani if args.piani in ("tutti", "primo") else (csv_arg(args.piani) or "tutti"),
        includi_non_diversificato=args.includi_non_diversificato,
        includi_altre_sedi=args.includi_altre_sedi,
        periodi=None if periodi is None else {norm_periodo(p) for p in periodi},
        scaglioni=not no_det, orari=not (no_det or args.no_orari),
        out=args.out, delay=args.delay, paralleli=args.paralleli, cache=args.cache)
    try:
        out = scarica(opt)
    except ValueError as e:
        sys.exit(str(e))
    if json.loads(out.read_text(encoding="utf-8"))["meta"].get("errore"):
        sys.exit(1)


if __name__ == "__main__":
    main()
