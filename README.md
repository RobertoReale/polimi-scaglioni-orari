# PoliMi – scaglioni e orari

**Cosa fa:** legge dal sito dei [Manifesti degli Studi del Politecnico di Milano](https://onlineservices.polimi.it/manifesti/manifesti/controller/ManifestoPublic.do?evn_DEFAULT=evento&lang=IT)
gli insegnamenti dei corsi di studio che scegli, con i loro **scaglioni** (gli studenti divisi per iniziale
del cognome, ognuno con i suoi docenti) e l'**orario delle lezioni** (giorni, ore, aule). Poi puoi
filtrare i dati e salvarli in **Excel, CSV, pagina web** o come **orario settimanale** da stampare.

**Come si usa, in due passi:**
1. **Scarica** i dati dei corsi che ti interessano (scheda 1).
2. **Esplora ed esporta**: consulta i dati, filtrali e salvali nel formato che preferisci (scheda 2).

Il programma legge soltanto il sito pubblico: non modifica nulla e non chiede credenziali.
Funziona su **Windows**, **Linux** e **macOS**.

Nella finestra, il pulsante **❓ Guida** spiega tutto questo e le parole da conoscere. Se tieni il mouse
fermo su un'opzione, compare una breve spiegazione di cosa fa.

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

## Parole da conoscere

| Parola | Significato |
|---|---|
| Corso di studi | per esempio *Ingegneria Informatica*; ogni corso ha un codice numerico (es. `531`) |
| Tipo di laurea | Laurea di primo livello, Laurea Magistrale, Ciclo Unico… *ord. 96/23* indica il regolamento (ordinamento) del corso |
| Piano di studi | una variante dello stesso corso (es. in italiano o in inglese, o in un'altra sede), con il suo elenco di insegnamenti. Il piano `***` raccoglie gli insegnamenti comuni a tutti i piani |
| Periodo didattico | 1° semestre, 2° semestre o annuale |
| Scaglione | gruppo di studenti di un insegnamento, diviso per iniziale del cognome, con i suoi docenti, orari e aule. `BRU – CON` vuol dire da BRU (compreso) a CON (escluso); `A – ZZZZ (unico)` vuol dire un solo gruppo per tutti |
| Fascia di cognomi | gli insegnamenti dividono i cognomi in modi diversi: la fascia è un intervallo in cui la divisione è la stessa per tutti, e serve a comporre l'orario completo di uno studente |

## Uso

### Scheda 1 · Scarica dati dal sito
1. **Anno accademico e sede**: le scelte arrivano direttamente dal sito. Con **Tipo di laurea** puoi mostrare
   solo i corsi di un tipo (per esempio *Laurea Magistrale*); *Tutti i tipi* li mostra tutti.
2. **Corsi di studio**: clic su un corso per selezionarlo. Un clic su una scuola o su un tipo di laurea
   (per esempio *Laurea di Primo Livello – ord. 96/23*) seleziona tutto il gruppo. Con **Cerca** trovi un
   corso per nome o codice.
3. **Cosa includere**:
   - anni di corso e periodo didattico;
   - piani di studio: tutti quelli della sede oppure solo il primo;
   - scaglioni e orario.

   Senza scaglioni e orario ottieni solo l'elenco degli insegnamenti, che si scarica molto più in fretta.
4. **Avvia scaricamento**. Prima di partire compare un **riepilogo delle scelte** da confermare. Una barra
   mostra l'avanzamento con una stima del tempo rimanente. Puoi interrompere in qualsiasi momento (anche
   chiudendo la finestra): i dati raccolti fino a quel punto vengono salvati comunque, e lo stesso vale se la
   connessione cade. Il risultato è un file `.json` nella cartella `output/`; alla fine un riepilogo dice
   quanti corsi, insegnamenti, scaglioni e lezioni sono stati letti.

Se manca una scelta, il programma lo segnala prima di partire. Nulla viene scelto al posto tuo.

### Scheda 2 · Esplora ed esporta
1. Scegli il **file di dati** da aprire.
2. Scegli la **tabella** da consultare:

   | Tabella | Una riga per… |
   |---|---|
   | Insegnamenti | ogni insegnamento di ogni piano: periodo, CFU, n. scaglioni, docenti |
   | Scaglioni | ogni scaglione con il suo orario settimanale: una colonna per giorno (Lunedì…Sabato, es. `15:15–18:15 aula 7.1.3`), le aule con l'edificio, il periodo delle lezioni, docenti, moduli |
   | Orario per cognome | ogni fascia di cognomi di un piano: l'orario settimanale completo, con **tutti** gli insegnamenti, di chi ha il cognome in quella fascia (es. `BRU – CON`: Analisi, Informatica e Geometria insieme) |
   | Lezioni (orario) | ogni lezione settimanale: giorno, ora di inizio e fine, aula ed edificio, date |
   | Corsi e piani | ogni piano di studio, compresi quelli scartati perché di un'altra sede |

   Perché "Scaglioni" ha più righe con le stesse lettere? Perché ogni insegnamento ha i suoi scaglioni:
   `BRU – CON` di Analisi e `BRU – CON` di Geometria sono due righe. **Orario per cognome** li mette insieme.
   Quando gli insegnamenti dividono i cognomi in modo diverso (uno A–M e M–Z, un altro A–E, E–P, P–Z), le
   fasce sono tutti gli intervalli tra i confini (A–E, E–M, M–P, P–Z), ognuna con lo scaglione giusto di ogni
   insegnamento.

   Negli insegnamenti annuali con orari diversi nei due semestri, accanto a ogni lezione della tabella
   Scaglioni ci sono le date in cui vale, es. `09:15–13:15 aula G.3 [24/02→26/05]`.

3. Restringi i risultati con i **filtri**: corso, piano, periodo, insegnamento, scaglione, giorno, aula,
   docente, oppure cerca un testo libero. I filtri si sommano, e ogni menu propone solo i valori ancora
   possibili con i filtri che lo precedono.
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
     per fascia di cognomi (l'orario completo di uno studente), per insegnamento, per piano, per aula o per docente. Si stampa o si salva in PDF dal browser.

   Le pagine web esportate usano sempre il tema chiaro, anche se il computer è in modalità scura.

## Cache
Le pagine scaricate restano salvate nella cartella `cache/`, così ripetere o ampliare uno scaricamento è
immediato. **Per avere i dati aggiornati dal sito usa "Svuota cache"** (scheda 1 → *Impostazioni avanzate…*),
oppure cancella la cartella `cache/`.

## Se qualcosa va storto
- **Sito non raggiungibile**: controlla la connessione e riavvia il programma.
- **Alcuni insegnamenti non letti**: li trovi segnalati nella colonna *Note*. Rilancia lo scaricamento con le
  stesse scelte: le pagine già lette vengono prese dalla cache, quindi si rileggono solo quelle mancanti.
- **"Non riesco a scrivere il file"**: il file è probabilmente aperto in Excel. Chiudilo e riprova.
- **Il sito risponde con errori**: aumenta la *Pausa tra le richieste* o diminuisci le *Richieste in parallelo*
  in *Impostazioni avanzate…*.

## Uso da terminale (facoltativo)
```bash
python scarica_manifesti.py --elenca          # elenca anni, sedi, scuole e codici dei corsi
python scarica_manifesti.py --sede MI --ordinamento 96/23 --anni-corso 1 --periodi annuale,1sem
python scarica_manifesti.py --corsi 531 --piani primo --no-orari --paralleli 2
python esporta.py output/FILE.json --tabella scaglioni --formato xlsx --filtro corso_codice=531
python esporta.py output/FILE.json --calendario scaglione --filtro piano_codice=IT1
python esporta.py output/FILE.json --calendario cognome --filtro corso_codice=531
python esporta.py output/FILE.json --tutte --formato xlsx
```
Ogni script mostra le opzioni disponibili con `--help`.

## Note
- Gli orari di inizio e fine vengono ricavati dalla griglia a quarti d'ora del sito.
- Alcuni insegnamenti non hanno orario: il sito indica "Non esistono occupazioni" oppure l'insegnamento
  è erogato da un ateneo partner.
- Il programma aspetta un attimo tra una richiesta e l'altra per non sovraccaricare il sito e ne fa al
  massimo 4 insieme, perché quasi tutto il tempo è attesa del server.
- Nel CSV i decimali usano la virgola, così Excel in italiano non li scambia per date.

## Per chi modifica il codice
Come funziona il programma, com'è fatto il sito, la struttura del file JSON e come fare le modifiche più
comuni: vedi **[SVILUPPO.md](SVILUPPO.md)**.
