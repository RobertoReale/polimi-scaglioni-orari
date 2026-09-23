# Guida per chi sviluppa

Come funziona il programma dall'interno. Per l'uso normale vedi il [README](README.md).

## In una frase

Il programma legge le pagine pubbliche dei Manifesti degli Studi del PoliMi, ne ricava corsi → piani →
insegnamenti → scaglioni → lezioni, salva tutto in un file JSON e da quel file costruisce tabelle da filtrare
ed esportare.

## I file e cosa fa ciascuno

| File | Responsabilità | Dipende da |
|---|---|---|
| `avvia.bat`, `avvia.sh` | trovano Python e lanciano `avvia.py` | — |
| `avvia.py` | controlla Python e tkinter, installa le librerie in `.venv` se mancano, apre l'interfaccia | `interfaccia.py` |
| `scarica_manifesti.py` | **legge il sito** e scrive il JSON. Nessuna interfaccia: si usa anche da terminale | `requests`, `bs4` |
| `esporta.py` | **dal JSON alle tabelle** e ai file (Excel, CSV, HTML, JSON, calendario). Non accede al sito | `openpyxl` |
| `interfaccia.py` | **solo l'interfaccia** (tkinter): raccoglie le scelte e chiama gli altri due moduli | i due moduli sopra |

La regola: la logica sta in `scarica_manifesti.py` ed `esporta.py`, che funzionano anche senza interfaccia;
`interfaccia.py` non decide niente sui dati.

## Percorso dei dati

```
 sito PoliMi ──► scarica_manifesti.scarica(Opzioni) ──► output/manifesti_<aa>_<sede>_<data>.json
                   │ pagine salvate in cache/                          │
                   ▼                                                   ▼
             (fase 1) corsi                                 esporta.tabelle(dati)
             (fase 2) anni → piani → insegnamenti              │  5 tabelle: liste di dict
             (fase 3) scaglioni + orario di ogni insegnamento  ▼
                                                            esporta.filtra(...)            filtri dell'utente
                                                               ▼
                                                            esporta.unisci_duplicati(...)  facoltativo
                                                               ▼
                                                            esporta.esporta(...) / esporta_xlsx(...) /
                                                            esporta_calendario(...)
```

## Com'è fatto il sito

Tutto passa da un'unica pagina, `ManifestoPublic.do`, che cambia secondo i parametri dell'indirizzo:

| Parametro | Significato | La pagina mostra |
|---|---|---|
| `aa` | anno accademico (`2026` = 2026/2027) | — |
| `sede` | `MI`, `BV`, `CO`, `CR`, `LC`, `MN`, `PC` oppure `ALL_SEDI` | — |
| `k_cf` | scuola | il menu dei corsi della scuola |
| `k_corso_la` | corso di studi | il menu degli anni e dei piani |
| `ac_ins` | anno di corso (`0` = tutti) | — |
| `k_indir` | piano di studio (`***` = offerta non diversificata) | la tabella degli insegnamenti |

I menu della pagina (`<select name="...">`) vengono letti con `select_options()`: è così che il programma
scopre anni accademici, sedi, scuole, corsi e piani disponibili, senza elenchi scritti nel codice.

Ogni insegnamento ha una pagina di dettaglio con una o più *sezioni*, ognuna con due schede caricate
a parte (richieste AJAX): **Dettaglio** (scaglioni, docenti, moduli) e **Orario didattico** (griglia
settimanale). Vedi `fetch_dettaglio()`.

**Punti fragili.** Se il sito cambia, è qui che il programma si rompe:
- le intestazioni in italiano delle tabelle: `MANIFESTO_KEYS`, `SCAGLIONI_KEYS` (per questo `lang` è
  sempre `IT`);
- le classi CSS: `TableDati`, `TitleInfoCard` (elenco insegnamenti), `tabs`, `ui-state-disabled`
  (schede del dettaglio), `SimpleCardOut`, `scrollTable`, `slot`, `data`, `dove`, `slotSem…` (orario);
- la griglia dell'orario: il sito non scrive l'ora di inizio e fine delle lezioni, la disegna. Ogni colonna
  è un quarto d'ora; `parse_orario_grid()` ricava l'ora dalla posizione e dalla larghezza della cella.

## Lo scaricamento: `scarica()` in tre fasi

1. **`_trova_corsi()`** – per ogni scuola, i corsi che rispettano i filtri (codici, ordinamento, tipo di laurea).
2. **`_elabora_corso()`** – per ogni corso (più corsi in parallelo) e ogni anno richiesto:
   - `_scegli_piani()` decide quali piani leggere: il piano `***` solo se richiesto o se è l'unico;
     un corso senza menu dei piani ha un piano unico;
   - `_piano_di_altra_sede()` scarta i piani di una città diversa da quella scelta (restano elencati in
     `piani_scartati`);
   - con l'opzione "primo" si tiene solo il primo piano **non scartato**;
   - gli insegnamenti vengono filtrati per periodo (`periodo_ok()`).
3. **`_scarica_dettagli()`** – per ogni insegnamento (più insegnamenti in parallelo) scaglioni e orario.

Regole che valgono ovunque:
- **Non si perdono dati.** Interruzione dell'utente, sito irraggiungibile, errore imprevisto: si salva
  quanto raccolto, con `meta.completo = false` e il motivo in `meta.interrotto` o `meta.errore`.
- **Un errore locale non ferma tutto.** Un corso o un insegnamento che non si legge riceve il campo `errore`
  e lo scaricamento continua; nelle tabelle compare nella colonna *Note* o *Stato*.
- **Cache.** `Client.get()` salva ogni pagina in `cache/<sha1 dell'indirizzo>.html` e la riusa. La scrittura
  è atomica (file temporaneo + rinomina), così un'interruzione non lascia pagine a metà.
- **Parallelismo.** `in_parallelo()` usa un pool di thread (normale: 4). `Client` tiene una sessione HTTP per
  thread. L'interruzione passa per un `threading.Event` (`stop`) che anche le pause controllano.

## Il file JSON

```
meta
  fonte, generato_il
  completo              false se interrotto o con errori
  interrotto / errore   il motivo, se non completo
  parametri             le Opzioni usate (+ aa e sede_nome effettivi)
  riepilogo             conteggi: corsi, piani, insegnamenti, scaglioni, lezioni, errori
corsi_di_studio[]
  scuola{codice, nome}, codice, nome, tipo_ordinamento, anni_disponibili
  errore                se il corso non si è potuto leggere
  note[]                es. "anno di corso 3 non disponibile"
  piani_scartati[]      codice, nome, sede, anno_corso  (piani di altre sedi)
  piani[]
    codice (null = piano unico), nome, sede, lingua, anno_corso, n_insegnamenti_totali
    insegnamenti[]
      codice, nome, ssd, tipo, periodo, cfu, lingua[], sede_erogazione, blocco, anno_corso,
      url_dettaglio, n_scaglioni        (più alcuni codici interni del sito)
      nota_dettaglio    perché non ci sono sezioni (es. erogato da un ateneo partner)
      errore            se il dettaglio non si è potuto leggere
      sezioni[]
        id_sezione, n_scaglioni
        scaglioni[]     da, a (iniziali del cognome: da compreso, a escluso), docenti[], righe[] (moduli)
        orario[]        scaglione_da, scaglione_a, giorno, giorno_n, inizio, fine, durata_min,
                        aula, aula_descrizione, attivita, dal, al, date_lezioni[], periodo_orario
        orario_nota     perché l'orario è vuoto
```

Lo stesso insegnamento compare in ogni piano (e corso) che lo contiene: nel JSON è ripetuto.

## Dal JSON alle tabelle (`esporta.py`)

`tabelle(dati)` "appiattisce" il JSON: ogni riga ripete il suo contesto (anno accademico, scuola, corso,
piano, anno), così si capisce da sola e si può filtrare per qualunque colonna. Una funzione per tipo di riga:

| Tabella | Costruita da | Una riga per |
|---|---|---|
| `insegnamenti` | `_riga_insegnamento()` | insegnamento di un piano |
| `scaglioni` | `_righe_scaglioni()` | scaglione, con le lezioni divise per giorno (`_giorni_scaglione()`) |
| `lezioni` | `_righe_lezioni()` | lezione settimanale |
| `cognomi` | `_righe_cognomi()` → `gruppi_cognomi()` | fascia di cognomi di un piano |
| `piani` | `_riga_piano()`, `_righe_piani_non_scaricati()` | piano (anche scartati o con errore) |

Le colonne di ogni tabella, e il loro ordine, sono in `COLONNE`; i nomi mostrati all'utente in `LABELS`;
le descrizioni in `DESCRIZIONI`.

Due logiche da conoscere:
- **`gruppi_cognomi()`** – ogni insegnamento divide i cognomi a modo suo. Le fasce sono gli intervalli
  tra *tutti* i confini degli scaglioni del piano; per ogni fascia si prende, da ogni insegnamento, lo
  scaglione che la contiene. Esempio nella docstring.
- **`unisci_duplicati()`** – due righe sono "la stessa" se coincidono in tutte le colonne tranne quelle di
  `VARIABILI_PER_PIANO` (corso, piano, link…); quelle colonne diventano elenchi separati da ` | `.

## L'interfaccia (`interfaccia.py`)

- `App` – la finestra: intestazione con la descrizione e il pulsante *Guida*, e due schede.
- `SchedaScarica` – le scelte diventano un oggetto `scarica_manifesti.Opzioni` in `_opzioni()`, che controlla
  anche che non manchi niente; `_riepilogo_scelte()` le mostra in parole prima di partire.
- `SchedaEsplora` – carica il JSON, costruisce le tabelle, applica i filtri, esporta.
- `Suggerimento` / `aiuto()` – la spiegazione che compare tenendo il mouse su un controllo.
- `GUIDA` – il testo della finestra *Guida*.

**Thread.** Lo scaricamento (e il caricamento dell'elenco dei corsi) gira in un thread separato, così la
finestra non si blocca. tkinter non permette di toccare i widget da un altro thread: il thread mette
messaggi nella coda `self.q` (`"log"`, `"prog"`, `"fine"`, `"finito"`), e `_svuota_coda()`, che gira ogni
100 ms nel thread della finestra, li trasforma in aggiornamenti.

## Principi da rispettare

- **Nessuna scelta al posto dell'utente.** I campi partono vuoti (`— scegli —`) e `_opzioni()` segnala cosa
  manca. Non aggiungere valori preimpostati nascosti.
- **Ogni opzione si spiega da sola.** Un controllo nuovo ha un'etichetta chiara e un `aiuto(...)`; se
  introduce un termine nuovo, aggiungilo anche in `GUIDA` e nel README.
- **Testi per l'utente in italiano semplice**, senza termini tecnici (niente "JSON parsing", "thread"…).
- **Non perdere dati** e **non fermare tutto per un errore locale** (vedi sopra).
- **Rispetto per il sito**: pausa tra le richieste e poche richieste in parallelo.

## Modifiche comuni

**Aggiungere una colonna a una tabella**
1. Calcola il valore nella funzione che costruisce quella riga (es. `_righe_scaglioni()`).
2. Aggiungi la chiave in `COLONNE[tabella]`, nella posizione in cui deve comparire.
3. Aggiungi il nome leggibile in `LABELS`.
4. Se la colonna cambia da un piano all'altro (come il corso), aggiungila a `VARIABILI_PER_PIANO`,
   altrimenti *unisci doppioni* non unirà più le righe.

**Aggiungere un filtro nella scheda 2**: una voce `(colonna, "Nome")` in `FILTRI` (`interfaccia.py`).

**Aggiungere una tabella**: una funzione `_righe_...()`, la chiamata in `tabelle()`, e le voci in `COLONNE`,
`DESCRIZIONI` e `NOMI_TABELLE` (`interfaccia.py`).

**Aggiungere un'opzione di scaricamento**: un campo in `Opzioni`, il suo uso in `scarica_manifesti.py`,
l'argomento da terminale in `main()`, il controllo in `SchedaScarica._costruisci()` (con `aiuto`), la lettura
in `_opzioni()` e una riga in `_riepilogo_scelte()`.

**Leggere un dato nuovo dal sito**: se è una colonna della tabella degli insegnamenti o degli scaglioni,
di solito basta aggiungere l'intestazione del sito in `MANIFESTO_KEYS` o `SCAGLIONI_KEYS`.

## Come verificare una modifica

Non ci sono test automatici. Il modo più affidabile è confrontare i risultati prima e dopo la modifica:

1. **Scaricamento** – con la cache piena è immediato:
   ```bash
   python scarica_manifesti.py --sede MI --corsi 531 --anni-corso 1 --piani primo --out prima.json
   # ... modifica ...
   python scarica_manifesti.py --sede MI --corsi 531 --anni-corso 1 --piani primo --out dopo.json
   ```
   I due file devono coincidere, a parte `meta.generato_il` e `meta.parametri.out`.
2. **Tabelle** – confronta `esporta.tabelle()` sugli stessi file in `output/` prima e dopo (per esempio
   salvando `json.dumps(tabelle(dati), sort_keys=True)`), anche con `unisci_duplicati()`.
3. **Interfaccia** – avvia `python avvia.py` e prova: scelte mancanti → messaggio; riepilogo prima
   dell'avvio; interruzione a metà → i dati parziali si aprono nella scheda 2; ogni esportazione si apre.
