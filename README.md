# PoliMi – scaglioni e orari

Programma per scaricare dal sito dei [Manifesti degli Studi del Politecnico di Milano](https://onlineservices.polimi.it/manifesti/manifesti/controller/ManifestoPublic.do?evn_DEFAULT=evento&lang=IT)
gli insegnamenti di uno o più corsi di studio, con i loro **scaglioni** (lettere, docenti, moduli) e
l'**orario delle lezioni** (giorni, ore, aule). I dati si possono poi filtrare e salvare in **Excel, CSV,
pagina web** o come **orario settimanale** da stampare.

Funziona su **Windows**, **Linux** e **macOS**.

## Installazione

1. Installa **Python 3.8 o più recente**:
   - Windows: scaricalo da [python.org](https://www.python.org/downloads/). Durante l'installazione spunta **"Add python.exe to PATH"**.
   - Linux (Ubuntu/Debian): `sudo apt install python3 python3-venv python3-tk`
2. Scarica questo progetto: pulsante verde **Code → Download ZIP**, poi estrai la cartella.

Le librerie necessarie (`requests`, `beautifulsoup4`, `openpyxl`) si installano da sole al primo avvio,
in una cartella locale `.venv`. Serve la connessione a internet.

## Avvio

| Sistema | Come |
|---|---|
| Windows | doppio clic su **`avvia.bat`** |
| Linux / macOS | da terminale nella cartella: **`bash avvia.sh`** |
| qualunque | `python avvia.py` |

## Uso

### Scheda 1 · Scarica dati dal sito
1. **Anno accademico e sede**: le scelte arrivano direttamente dal sito. Con **Tipo di laurea** puoi mostrare
   solo i corsi di un tipo (per esempio *Laurea Magistrale* o *Laurea di Primo Livello*); *Tutti i tipi* li mostra tutti.
2. **Corsi di studio**: clic su un corso per selezionarlo. Un clic su una scuola o su un tipo di laurea
   (per esempio *Laurea di Primo Livello – ord. 96/23*) seleziona tutto il gruppo. Con **Cerca** filtri l'elenco.
3. **Cosa includere**:
   - anni di corso e periodo didattico;
   - piani di studio: tutti quelli della sede oppure solo il primo;
   - scaglioni e orario.

   Senza scaglioni e orario ottieni solo l'elenco degli insegnamenti, che si scarica molto più in fretta.
4. **Avvia scaricamento**. Una barra mostra l'avanzamento con una stima del tempo rimanente. Puoi interrompere
   in qualsiasi momento (anche chiudendo la finestra): i dati raccolti fino a quel punto vengono salvati comunque,
   e lo stesso vale se la connessione cade. Il risultato è un file `.json` nella cartella `output/`; alla fine
   un riepilogo dice quanti corsi, insegnamenti, scaglioni e lezioni sono stati letti.

Se manca una scelta, il programma lo segnala prima di partire. Nulla viene scelto al posto tuo.

### Scheda 2 · Esplora ed esporta
1. Scegli il **file di dati** da aprire.
2. Scegli la **tabella** da consultare:

   | Tabella | Una riga per… |
   |---|---|
   | Insegnamenti | ogni insegnamento di ogni piano: periodo, CFU, n. scaglioni, docenti |
   | Scaglioni | ogni scaglione con il suo orario settimanale: una colonna per giorno (Lunedì…Sabato, es. `15:15–18:15 aula 7.1.3`), le aule con l'edificio, il periodo delle lezioni, docenti, moduli |
   | Lezioni (orario) | ogni lezione settimanale: giorno, ora di inizio e fine, aula ed edificio, date |
   | Corsi e piani | ogni piano di studio, compresi quelli scartati perché di un'altra sede |

   Negli insegnamenti annuali con orari diversi nei due semestri, accanto a ogni lezione della tabella
   Scaglioni ci sono le date in cui vale, es. `09:15–13:15 aula G.3 [24/02→26/05]`.

3. Restringi i risultati con i **filtri**: corso, piano, periodo, insegnamento, scaglione, giorno, aula,
   docente, oppure cerca un testo libero.
   - L'opzione **"Una riga sola per ciò che si ripete in più corsi"** elimina i doppioni: uno stesso
     insegnamento (o scaglione, o lezione) può comparire in più corsi e piani di studio. Resta una riga sola,
     e le colonne Corso e Piano elencano tutti i corsi separati da `|`. Per esempio, nella tabella Scaglioni
     Analisi Matematica 1 passa da 32 righe (8 scaglioni × 4 corsi) a 8. Vale anche per i file esportati.
   - Con **Colonne…** scegli quali colonne tenere.
   - Con un doppio clic su una riga la vedi per intero.
4. **Salva** le righe filtrate:
   - **Excel**, **CSV** (si apre direttamente in Excel), **pagina web** con ricerca e ordinamento, **JSON**;
   - **Excel con tutte le tabelle**: un foglio per ogni tabella, con gli stessi filtri applicati;
   - **Orario settimanale**: una pagina con le griglie Lunedì–Venerdì. Puoi raggrupparle per scaglione,
     per insegnamento, per piano, per aula o per docente. Si stampa o si salva in PDF dal browser.

   Le pagine web esportate usano sempre il tema chiaro, anche se il computer è in modalità scura.

## Cache
Le pagine scaricate restano salvate nella cartella `cache/`, così ripetere o ampliare uno scaricamento è
immediato. **Per avere i dati aggiornati dal sito usa "Svuota cache"** nella scheda 1, oppure cancella la cartella `cache/`.

## Uso da terminale (facoltativo)
```bash
python scarica_manifesti.py --elenca          # elenca anni, sedi, scuole e codici dei corsi
python scarica_manifesti.py --sede MI --ordinamento 96/23 --anni-corso 1 --periodi annuale,1sem
python scarica_manifesti.py --corsi 531 --piani primo --no-orari --paralleli 2
python esporta.py output/FILE.json --tabella scaglioni --formato xlsx --filtro corso_codice=531
python esporta.py output/FILE.json --calendario scaglione --filtro piano_codice=IT1
python esporta.py output/FILE.json --tutte --formato xlsx
```
Ogni script mostra le opzioni disponibili con `--help`.

## File del progetto
| File | Contenuto |
|---|---|
| `avvia.bat`, `avvia.sh`, `avvia.py` | avvio: controllano e installano le librerie, poi aprono l'interfaccia |
| `interfaccia.py` | interfaccia grafica (tkinter) |
| `scarica_manifesti.py` | lettura del sito: corsi, piani, insegnamenti, scaglioni, orari → JSON |
| `esporta.py` | trasformazione del JSON in tabelle ed esportazione (Excel, CSV, HTML, calendario) |

## Struttura del file JSON
```
meta                      parametri usati, data, completo (false se interrotto)
corsi_di_studio[]
  scuola, codice, nome, tipo_ordinamento, anni_disponibili
  piani_scartati[]        piani di altre sedi
  piani[]
    codice, nome, sede, lingua, anno_corso
    insegnamenti[]
      codice, nome, ssd, tipo, periodo, cfu, lingua, sede_erogazione, blocco, url_dettaglio
      n_scaglioni, nota_dettaglio (es. corsi erogati da atenei partner)
      sezioni[]
        scaglioni[]       da, a (iniziali del cognome), docenti, righe[] (moduli)
        orario[]          scaglione_da/a, giorno, inizio, fine, aula, aula_descrizione,
                          attivita, dal, al, date_lezioni, periodo_orario
```

## Note
- Gli orari di inizio e fine vengono ricavati dalla griglia a quarti d'ora del sito.
- Alcuni insegnamenti non hanno orario: il sito indica "Non esistono occupazioni" oppure l'insegnamento
  è erogato da un ateneo partner.
- Il programma aspetta un attimo tra una richiesta e l'altra per non sovraccaricare il sito e ne fa al
  massimo 4 insieme, perché quasi tutto il tempo è attesa del server. Pausa e richieste in parallelo si
  regolano in *Avanzate*.
- Nelle tabelle, lo scaglione `A – ZZZZ (unico)` vuol dire che c'è un solo scaglione per tutti gli studenti.
- Nel CSV i decimali usano la virgola, così Excel in italiano non li scambia per date.
