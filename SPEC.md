# Prompt per Claude Code — Portale generazione schede tecniche Trafo Elettro

Copia tutto questo blocco in Claude Code come istruzione iniziale del progetto.

---

## Obiettivo

Web app interna e condivisa che, dato un file Excel di estrazione dati prodotto dall'ufficio tecnico di Trafo Elettro, genera un PDF "scheda tecnica" del trasformatore, brandizzato e curato graficamente. L'app valida i campi obbligatori, e quando mancano dati o trova valori sospetti li chiede all'utente in un form dinamico prima di produrre il PDF.

## Vincoli architetturali (non derogabili)

- **Nessun LLM a runtime.** Il parsing è 100% deterministico. L'Excel ha schema fisso, key-value. Niente chiamate API per leggere o interpretare i valori.
- **Stateless.** Genera-e-scarica. Nessun database, nessun object storage, nessun archivio storico, nessun invio email.
- **Stack bloccato:** Python 3.11+, Streamlit (UI + upload + form + download), openpyxl (lettura Excel), Jinja2 + WeasyPrint (HTML/CSS → PDF), Docker, deploy su Render.
- **Auth a password** a livello app (singola password condivisa via variabile d'ambiente), perché l'URL è raggiungibile in rete.

## Lingua e formato dei dati

- **Il PDF è sempre e interamente in inglese.** Tutte le etichette, le sezioni, le intestazioni e anche i **valori testuali** vanno in inglese. I dati arrivano dall'Excel con valori enumerati in italiano (es. `Conservatore`, `Reg. sottocarico`, `Strati`, `Radiatori`): vanno tradotti con una mappa di valori controllata (`value_map` nello YAML, vedi sotto). La UI interna del portale (form, messaggi) resta pure in italiano per velocità degli operatori; cambia solo il PDF.
- **Formato numerico coerente col dato.** Numeri in formato inglese (punto decimale, virgola per le migliaia). Ogni campo è formattato secondo la sua unità e i suoi decimali definiti nello YAML. Esempi: `0.0125` come `% gradino` → `1.25 %`; `20000` V → `20,000 V`; perdite e pesi interi; mai stampare `0.13319016617797527`. Il valore stampato deve essere coerente con come quel dato si esprime in una scheda tecnica reale.

## Formato di input (fisso, concordato con l'ufficio tecnico)

- Un solo foglio chiamato `Foglio1`.
- Due colonne: A = etichetta (italiano), B = valore. Riga 1 = intestazione (`Descrizione` / `Valore`).
- 72 parametri nelle righe 2–73, sempre nello stesso ordine e con le stesse etichette.
- Le etichette in colonna A sono la chiave di lookup. Leggi per etichetta, non per coordinata fissa, così resti robusto a piccoli scostamenti di riga.

### Le 72 etichette esatte (colonna A, righe 2–73)

```
Codifica foglio calcolo
Codice prodotto
Serie
Tipo casa
Tipo sistema raffreddamento
Raffreddamento
Temp. Ambiente MIN
Temp. Ambiente MAX
Altitudine
THDi
Potenza nominale
Potenza reg. forzato
Potenza reattiva
Avvolg. stabilizzatore? Per trafo EAR
Avvolg. ausiliario? Per trafo EAR
Regolazione? Per trafo EAR
Corrente guasto omopolare
Durata guasto omopolare
Corrente guasto permanente
Numero fasi
Frequenza
Tensione MT1
Tensione MT2
Tipo commutatore
Posizioni + rif. MT1
Posizioni - rif. MT1
% gradino rif. MT1
Tensione BT
Collegamento MT
Collegamento BT
Collegamento TER
Indice orario
Gruppo vettoriale
Perdite a vuoto
Perdite a carico 75°C
Perdite a carico 120°C
Corrente a vuoto %
Impedenza di cortocircuito %
PEI / MEPS / HEPS in AN/ONAN
Classe termica MT
Classe termica BT
Classe isolamento MT
Classe isolamento BT
Materiale MT
Materiale BT
Materiale ST
Tipo avvolg. MT
Tipo avvolg. BT
Tipo avvolg. ST
Sovratemperatura olio
Sovratemperatura avvolg. MT
Sovratemperatura avvolg. BT
LpA
LwA
Zo
Ro
Xo
Classe amb / clim / fuoco
Lunghezza trafo
Larghezza trafo
Altezza trafo
Peso totale maggiorato
Peso olio maggiorato
Tipologia olio
Interasse ruote
Ø ruote
K inserzione
Costante di tempo (sec)
Tempo emivalore (msec)
Induzione nucleo
Densità corrente MT (Sn/Pos Nom)
Densità corrente BT (Sn/Pos Nom)
```

## Campi liberi raccolti nel portale (non sono nell'Excel)

- `Cliente` (obbligatorio, testo)
- `Progetto` (obbligatorio, testo)
- `Note` (facoltativo, testo lungo, va nella sezione finale della scheda)
- `Offer #` e `Pos.` (facoltativi, testo, vanno nell'header della scheda)

## Schema canonico (`schema.yaml`)

Crea un file `parser/schema.yaml` che è l'unica fonte di verità del mapping. Per ogni campo definisci: etichetta italiana (chiave di lookup), nome inglese da stampare in scheda, sezione, unità di misura, n. decimali per l'arrotondamento, famiglia di applicabilità (`both` / `olio` / `resina`), flag `include_in_sheet`, flag `required`, e `suppress_if_zero`.

Mappatura di default da implementare (modificabile in seguito editando solo lo YAML):

| Etichetta IT | Label EN | Sezione | Unità | Famiglia | In scheda | Obblig. |
|---|---|---|---|---|---|---|
| Codice prodotto | Product code | Header | — | both | sì | no |
| Serie | Type / Series | General | — | both | sì | sì |
| Tipo casa | Construction | General | — | olio | sì | no |
| Tipo sistema raffreddamento | Cooling system | General | — | both | sì | no |
| Raffreddamento | Cooling | General | — | both | sì | sì |
| THDi | Loading application (THDi) | General | % | both | sì | no |
| Tipologia olio | Oil type | General | — | olio | sì | sì (olio) |
| Classe amb / clim / fuoco | Environmental / climatic / fire class | General | — | resina | sì | sì (resina) |
| Numero fasi | Number of phases | Electrical | — | both | sì | sì |
| Frequenza | Frequency | Electrical | Hz | both | sì | sì |
| Gruppo vettoriale | Vector group | Electrical | — | both | sì | sì |
| Indice orario | Clock index | Electrical | — | both | sì | no |
| Collegamento MT | HV connection | Electrical | — | both | sì | no |
| Collegamento BT | LV connection | Electrical | — | both | sì | no |
| Collegamento TER | TER connection | Electrical | — | both | sì | no |
| Perdite a vuoto | No-load losses | Electrical | W | both | sì | sì |
| Perdite a carico 75°C | Load losses 75°C | Electrical | W | olio | sì | sì (olio) |
| Perdite a carico 120°C | Load losses 120°C | Electrical | W | resina | sì | sì (resina) |
| Corrente a vuoto % | No-load current | Electrical | % | both | sì | no |
| Impedenza di cortocircuito % | Short-circuit impedance | Electrical | % | both | sì | sì |
| PEI / MEPS / HEPS in AN/ONAN | Efficiency index (PEI) | Electrical | % | both | sì | no |
| Potenza nominale | Rated power | Ratings | kVA | both | sì | sì |
| Potenza reg. forzato | Forced-cooling power | Ratings | kVA | olio | sì | no (suppress_if_zero) |
| Potenza reattiva | Reactive power | Ratings | kVAr | both | sì | no (suppress_if_zero) |
| Tensione MT1 | HV voltage | Ratings | V | both | sì | sì |
| Tensione MT2 | HV voltage (2nd) | Ratings | V | both | sì | no (suppress_if_zero) |
| Tensione BT | LV voltage | Ratings | V | both | sì | sì |
| Tipo commutatore | Tap changer | Ratings | — | both | sì | no |
| Posizioni + rif. MT1 | Tap positions (+) | Ratings | — | both | sì | no |
| Posizioni - rif. MT1 | Tap positions (−) | Ratings | — | both | sì | no |
| % gradino rif. MT1 | Step per tap | Ratings | % | both | sì | no |
| Classe isolamento MT | HV insulation level | Ratings | kV | both | sì | sì |
| Classe isolamento BT | LV insulation level | Ratings | kV | both | sì | sì |
| Materiale MT | HV winding material | Ratings | — | both | sì | no |
| Materiale BT | LV winding material | Ratings | — | both | sì | no |
| Materiale ST | TER winding material | Ratings | — | both | sì | no |
| Tipo avvolg. MT | HV winding type | Ratings | — | both | sì | no |
| Tipo avvolg. BT | LV winding type | Ratings | — | both | sì | no |
| Tipo avvolg. ST | TER winding type | Ratings | — | both | sì | no |
| Classe termica MT | HV thermal class | Ratings | — | both | sì | no |
| Classe termica BT | LV thermal class | Ratings | — | both | sì | no |
| Temp. Ambiente MIN | Ambient temp. min | Environmental | °C | both | sì | no |
| Temp. Ambiente MAX | Ambient temp. max | Environmental | °C | both | sì | no |
| Altitudine | Max installation altitude | Environmental | m a.s.l. | both | sì | no |
| Sovratemperatura olio | Oil temperature rise | Environmental | °C | olio | sì | no |
| Sovratemperatura avvolg. MT | HV winding temp. rise | Environmental | °C | both | sì | no |
| Sovratemperatura avvolg. BT | LV winding temp. rise | Environmental | °C | both | sì | no |
| LpA | Sound pressure level | Environmental | dBA | both | sì | no |
| LwA | Sound power level | Environmental | dBA | both | sì | no |
| Lunghezza trafo | Length | Dimensions & weight | mm | both | sì | no |
| Larghezza trafo | Width | Dimensions & weight | mm | both | sì | no |
| Altezza trafo | Height | Dimensions & weight | mm | both | sì | no |
| Peso totale maggiorato | Total weight | Dimensions & weight | kg | both | sì | no |
| Peso olio maggiorato | Oil weight | Dimensions & weight | kg | olio | sì | no |
| Interasse ruote | Wheel gauge | Dimensions & weight | mm | both | sì | no |
| Ø ruote | Wheel diameter | Dimensions & weight | mm | both | sì | no |

Campi da **escludere** dalla scheda cliente per default (parametri di calcolo interni; restano nello YAML con `include_in_sheet: false`, così basta un flag per riattivarli): `Codifica foglio calcolo`, `Potenza reattiva` se non serve, `Avvolg. stabilizzatore/ausiliario/Regolazione (EAR)`, `Corrente/Durata guasto omopolare`, `Corrente guasto permanente`, `Zo`, `Ro`, `Xo`, `K inserzione`, `Costante di tempo (sec)`, `Tempo emivalore (msec)`, `Induzione nucleo`, `Densità corrente MT/BT`.

### Traduzione dei valori testuali (`value_map`)

Per i campi con valori enumerati in italiano, definisci nello YAML una `value_map` italiano→inglese. Set di partenza da implementare (estendibile):

```yaml
value_map:
  Tipo casa:
    Conservatore: Conservator
    Ermetico: Hermetically sealed
  Tipo sistema raffreddamento:
    Radiatori: Radiators
    Onde: Corrugated walls
  Tipo commutatore:
    Reg. sottocarico: On-load tap changer (OLTC)
    A vuoto: Off-circuit tap changer (DETC)
    Reg. a vuoto: Off-circuit tap changer (DETC)
    Nessuno: None
  Tipo avvolg. MT: &winding
    Strati: Layer
    Dischi: Disc
    Continuo: Continuous disc
    Elicoidale: Helical
    Bobine: Coil
  Tipo avvolg. BT: *winding
  Tipo avvolg. ST: *winding
  Materiale MT: &material
    Al: Aluminium
    Cu: Copper
  Materiale BT: *material
  Materiale ST: *material
```

Codici che restano invariati (standard internazionali, non tradurre): `Raffreddamento` (ONAN, KNAN, AN…), `Gruppo vettoriale` (Dyn11…), `Collegamento MT/BT` (D, y, z…), `Serie`, `Tipologia olio` (nomi commerciali: MIDEL EN, TRANSAG 10LB…).

**Regola di sicurezza:** se un valore testuale non è nella `value_map` e il campo lo richiede, **non** stamparlo in italiano. Segnalalo nel form dinamico come valore da tradurre/confermare. Questo aggancia la traduzione alla stessa rete di validazione dei dati mancanti.

## Derivazione della famiglia (olio vs resina)

Non c'è un campo esplicito. Deriva così, in ordine:
1. Se `Tipologia olio` è valorizzato → **olio**.
2. Altrimenti se `Raffreddamento` ∈ {`AN`, `AF`, `ANAF`} (raffreddamento ad aria) → **resina**.
3. Altrimenti se `Raffreddamento` inizia per `O` o `K` (es. `ONAN`, `KNAN`) → **olio**.
4. Fallback: se `Serie` contiene `-R` → resina, altrimenti olio.

La famiglia decide quali campi obbligatori applicare e quali sezioni/righe mostrare.

## Normalizzazione dei tipi (regole esplicite)

1. **Bug date.** Excel può convertire i campi `Classe isolamento MT/BT` (formato tipo `24 / 50 / 125`) in una data. Se leggendo uno di questi campi trovi un `datetime`, **non** tentare di indovinare il valore: marcalo come *valore sospetto* e chiedilo nel form dinamico. In parallelo, nota nel README che nel template Excel standardizzato quelle celle vanno formattate come **Testo** all'origine per prevenire la conversione.
2. **Arrotondamenti.** Applica i decimali definiti per campo nello YAML. Default: perdite e pesi = interi; tensioni e dimensioni = interi; percentuali = max 2 decimali; `% gradino` va formattato come percentuale (`0.0125` → `1,25 %`); temperature = interi. Mai stampare numeri tipo `0.13319016617797527`.
3. **None vs 0.** `None` = dato non fornito. `0` = mostra solo se ha senso; per i campi con `suppress_if_zero: true` (es. `Tensione MT2`, `Potenza reg. forzato`, `Potenza reattiva`) ometti la riga se il valore è 0 o None.
4. **Soppressione righe.** Una riga non compare nel PDF se: valore None, oppure famiglia non applicabile, oppure `suppress_if_zero` e valore 0.

## Validazione e form dinamico (UX) — facilitare il più possibile

1. Dopo il parse, l'app calcola tre liste: **(a)** campi `required` mancanti per la famiglia derivata + `Cliente` e `Progetto`; **(b)** valori sospetti (bug date sui campi `Classe isolamento`); **(c)** valori testuali non presenti nella `value_map` da tradurre/confermare.
2. Se le tre liste sono vuote → vai diretto al riepilogo e al pulsante di generazione.
3. Altrimenti mostra un form Streamlit con **solo** gli elementi da gestire, raggruppati per sezione, ognuno con: etichetta in italiano + inglese, unità, e per i sospetti il **valore grezzo letto da Excel** così l'utente capisce cosa correggere.
4. **Ogni campo ha accanto un'opzione "Ometti (non applicabile)".** Tre esiti per campo:
   - l'utente inserisce il valore → usato;
   - l'utente spunta "Ometti" → la riga **non compare** nel PDF e l'obbligo è considerato soddisfatto (omissione volontaria, registrata in sessione);
   - resta vuoto e non omesso → blocca la generazione.
5. **Facilitazioni richieste** (questa fase deve costare il minimo sforzo):
   - "Ometti" è un solo clic per campo, più un pulsante "Ometti tutti i rimanenti" per chiudere in fretta i campi non pertinenti;
   - i campi pre-compilati col valore letto da Excel restano modificabili;
   - default sensati per unità e tipo di input (numerico vs testo) presi dallo YAML;
   - stato persistito in `st.session_state`, così tornando indietro non si perde nulla;
   - il pulsante "Genera PDF" è abilitato solo quando ogni obbligatorio è valorizzato **o** omesso.
6. Prima della generazione, mostra **sempre** il riepilogo completo dei valori che finiranno in scheda (estratti + inseriti + tradotti), con marcatura visibile su quelli inseriti a mano o tradotti. Nessun valore arriva al PDF senza passare da qui.

La UI del form è in italiano (operatori interni). Il PDF resta interamente in inglese.

## Immagine del trasformatore e quote dimensionali

La scheda mostra un'immagine del trasformatore scelta automaticamente in base al tipo derivato dall'Excel, con sopra le tre quote d'ingombro.

### Le 4 immagini (in `assets/transformers/`)

| File | Tipo rappresentato | Origine |
|---|---|---|
| `oil_conservator_radiators.png` | Olio, conservatore, **banchi radiatori separati** + commutatore sottocarico | Image 1 |
| `oil_conservator_corrugated.png` | Olio, conservatore, **pareti corrugate (onde)** | Image 2 |
| `oil_hermetic.png` | Olio, **ermetico** | Image 3 |
| `resin_open.png` | **Resina** a colonne | Image 5 |

(L'immagine cabina/enclosure è accantonata per ora. Si riaggiunge in futuro introducendo un campo `Cabina` nell'Excel standardizzato.)

### Regola di selezione (deterministica dai dati)

```
famiglia == resina                              -> resin_open
famiglia == olio:
    Tipo casa == Ermetico                       -> oil_hermetic
    Tipo casa == Conservatore:
        Tipo sistema raffreddamento == Radiatori -> oil_conservator_radiators
        Tipo sistema raffreddamento == Onde      -> oil_conservator_corrugated
```

Verificata sui 5 file di test, risolve tutti in modo univoco: 1250kVA → radiators; EOT888 → corrugated; EOT999 e TOT999 (Ermetico) → hermetic; TRT999 → resin_open.

**Safety net:** se `Tipo casa` o `Tipo sistema raffreddamento` contiene un valore non previsto, l'app non indovina: chiede l'immagine nel form con un menù a tendina sulle 4 opzioni.

### Quote dimensionali sull'immagine

Sotto l'immagine, una didascalia sempre corretta:
`Overall dimensions (L × W × H): {Lunghezza} × {Larghezza} × {Altezza} mm`
dove L = `Lunghezza trafo`, W = `Larghezza trafo`, H = `Altezza trafo`. Valori interi, separatore migliaia inglese.

In più, callout opzionali posizionati sull'immagine (L verso lo spigolo frontale-basso, W verso il lato, H verso lo spigolo verticale). Dato che ogni immagine ha prospettiva diversa, le posizioni dei callout sono **ancore configurabili per immagine** (coordinate normalizzate 0–1) in un file `assets/transformers/anchors.yaml`, così si registrano una volta per ciascuna immagine e restano stabili. Implementa i callout come overlay (HTML posizionato o SVG sopra l'immagine) compatibile con WeasyPrint. Se le ancore non sono definite per un'immagine, mostra solo la didascalia.



## Output PDF — condizioni di stile (precise)

Il PDF deve essere percepito come un documento Trafo Elettro, stessa famiglia visiva del Company Profile. Densità maggiore (è una scheda dati), ma stessi font, palette, motivo diagonale e impaginazione editoriale. Formato **A4 verticale**, multipagina, WeasyPrint.

### Font (Zalando Sans, self-hostato)

Niente caricamento da Google Fonts a runtime: WeasyPrint non lo gestisce in modo affidabile e l'output non sarebbe riproducibile. Scarica i TTF statici e dichiarali con `@font-face` puntando a `assets/fonts/`.

- Fonte: GitHub `zalando/sans` (release) oppure Fontsource (`@fontsource/zalando-sans`), licenza OFL.
- Pesi da includere: **300 Light, 400 Regular, 500 Medium, 600 SemiBold, 700 Bold**.
- Famiglia corpo/tabelle: **Zalando Sans**.
- Famiglia titoli display: **Zalando Sans SemiExpanded** (pesi 400/600), per avvicinarsi ai titoloni larghi del profilo. Se preferisci una sola famiglia, usa Zalando Sans ovunque: rendi il nome famiglia una variabile CSS così si cambia in un punto.
- Fallback tecnico solo per non rompere il layout in dev: `Arial, sans-serif`.

### Design token (variabili CSS, definirle in `:root`)

```css
--ink:        #1A1A1A;   /* testo principale */
--ink-soft:   #595959;   /* testo secondario, unità, label */
--line:       #D4D6D9;   /* hairline tabelle */
--rule:       #1A1A1A;   /* righe forti sotto i titoli sezione */
--panel:      #ECEDEF;   /* fasce/intestazioni tabella grigio chiaro */
--panel-dark: #111111;   /* fascia nera opzionale (come pag. 4 del profilo) */
--accent:     #F2C200;   /* giallo segnale, UNICO colore acceso. Campiona l'esatto dal logo */
--white:      #FFFFFF;
--diagonal:   112deg;    /* angolo del motivo diagonale, coerente in tutto il doc */
```

Regola d'uso del colore: bianco e near-black dominano. Il giallo è un accento puntuale (tick dei titoli sezione, filetto dell'header, slash decorativi), mai fondali estesi né testo. Niente altri colori.

### Scala tipografica (in punti, documento di stampa)

| Elemento | Famiglia / peso | Dim. | Note |
|---|---|---|---|
| Titolo documento ("Technical Datasheet") | SemiExpanded 400 | 30 pt | tracking −0.01em, leading stretto |
| Designazione unità (Serie + potenza) | SemiExpanded 600 | 15 pt | sotto il titolo |
| Eyebrow sezione (GENERAL, ELECTRICAL…) | Sans 700 | 8.5 pt | maiuscolo, tracking 0.14em, con filetto + tick giallo |
| Label parametro (colonna sinistra) | Sans 400 | 9 pt | colore --ink |
| Valore (colonna destra) | Sans 500 | 9 pt | --ink |
| Unità (dopo il valore) | Sans 400 | 8 pt | --ink-soft |
| Label meta (Client, Project…) | Sans 600 | 7 pt | maiuscolo, tracking 0.12em, --ink-soft |
| Valore meta | Sans 500 | 11 pt | --ink |
| Note / corpo | Sans 400 | 9 pt | line-height 1.45 |
| Didascalia quote | Sans 500 | 8.5 pt | |
| Footer | Sans 400 | 7 pt | --ink-soft |

### Griglia e spaziatura

- Margini `@page`: top 22mm (fascia header), bottom 16mm (footer), left/right 15mm.
- Scala spazi: 4 / 8 / 12 / 16 / 24 / 32 pt. Usare solo questi valori.
- Filetti: hairline tabella 0.5pt `--line`; filetto sotto eyebrow 1pt `--rule`; tick giallo 2.5pt × 12pt `--accent` a sinistra dell'eyebrow.
- Whitespace generoso tra sezioni (24–32pt). La scheda respira come il profilo, non è un modulo fitto.

### Header e footer ricorrenti (running elements WeasyPrint)

- Header (`position: running()` + `@page { @top-* }` o elemento posizionato): logo a sinistra (SVG), a destra la tagline "endless transformation" in 7pt tracking 0.2em --ink-soft. Sotto, un **filetto diagonale** sottile a tutta larghezza con l'angolo `--diagonal` (slash che richiama il profilo), realizzato in SVG inline o con `linear-gradient` hard-stop, non con `clip-path`.
- Footer: `TRAFO ELETTRO SRL · Via Ponte Poscola · 36075 Montecchio Maggiore (VI) · trafoelettro.com` a sinistra, `Page counter(page) / counter(pages)` a destra, 7pt --ink-soft, filetto hairline sopra.

### Blocco intestazione (prima pagina)

- Titolo "Technical Datasheet" grande, allineato a sinistra.
- Riga designazione unità sotto.
- Blocco meta a destra o sotto: Client, Project, Date, Product code, Offer #, Pos. Coppie label/valore con label-eyebrow sopra il valore.
- Un blocco grafico diagonale (fascia `--panel` tagliata a `--diagonal`) come ancora visiva dell'header, sul modello dei tagli del profilo.

### Titolo di sezione

Eyebrow maiuscolo (tick giallo a sinistra + filetto `--rule` di 1pt sotto, lungo quanto la colonna). Spaziatura 0.14em. Niente fondino colorato.

### Tabelle dati

- Due colonne: parametro (sinistra, ~58%) e valore+unità (destra). Valore e unità nella stessa cella, unità in `--ink-soft`.
- Righe separate da hairline 0.5pt `--line`. Niente bordi verticali. Zebra molto leggera opzionale (`--panel` al 40%): se la usi, tienila quasi impercettibile.
- Intestazione di gruppo (dove serve) su fascia `--panel`.
- **Sezione Ratings**: layout a due colonne affiancate **HV / LV** (MT/BT), con intestazione HV e LV, così i parametri di primario e secondario si leggono in parallelo come in una scheda tecnica reale.
- Le righe assenti (valore None / non applicabile / omesse) non compaiono: niente celle vuote o trattini.
- `page-break-inside: avoid` sui blocchi sezione per non spezzare una tabella a metà.

### Figura trasformatore

- Immagine selezionata dalla regola, contenuta in una `figure` con padding, su fondo bianco.
- Didascalia quote sotto (vedi sezione dedicata). Callout opzionali in overlay SVG con ancore configurabili.
- Se vuoi un tocco di brand, incornicia la figura con un angolo tagliato a `--diagonal` (solo decorativo, non sull'immagine tecnica che resta integra e leggibile).

### Note WeasyPrint (per non perdere tempo)

- Diagonali: usa **SVG inline** o `linear-gradient` con hard-stop a `--diagonal`. Evita `clip-path` (supporto parziale).
- Header/footer ripetuti: `position: running()` + `content: element()`, oppure i margin box `@top-left/@bottom-*`. Numerazione con `counter(page)`/`counter(pages)`.
- `@font-face` con i TTF locali; verifica che i pesi richiamati esistano tra quelli scaricati (300/400/500/600/700), altrimenti WeasyPrint fa fallback silenzioso e il risultato cambia.
- Testa il rendering dentro il container Docker, non solo in locale: i font e le librerie di sistema devono essere gli stessi, o l'output diverge.

### Asset che fornisco io (prevedi i percorsi)

Logo Trafo Elettro e segno fulmine in SVG ad alta risoluzione in `assets/`. Le 4 immagini trasformatore in `assets/transformers/`. I TTF di Zalando Sans in `assets/fonts/` (scaricabili dalle fonti sopra).

## Auth

Gate a password singola condivisa, letta da variabile d'ambiente `APP_PASSWORD`. All'avvio, se la sessione non è autenticata, mostra solo un campo password; alla corrispondenza sblocca l'app e salva il flag in `st.session_state`. Predisponi il codice per passare in futuro a utenti multipli senza riscrivere il flusso.

## Docker + Render

`Dockerfile` da `python:3.11-slim`. Installa via apt le librerie di sistema richieste da WeasyPrint: `libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf-2.0-0 libcairo2 libffi-dev shared-mime-info`. `pip install` da `requirements.txt` (`streamlit`, `weasyprint`, `jinja2`, `openpyxl`, `pyyaml`, `pandas`). Copia `assets/` e il codice. `EXPOSE 8501`. `CMD ["streamlit","run","app.py","--server.port=8501","--server.address=0.0.0.0"]`.

Render: servizio Docker, nessun persistent disk (app stateless), variabile d'ambiente `APP_PASSWORD`, health check sulla porta 8501. Aggiungi un `render.yaml` opzionale.

## Struttura cartelle attesa

```
trafo-datasheet/
  app.py                    # Streamlit: auth, upload, validazione, form dinamico, riepilogo, genera
  parser/
    schema.yaml             # mapping canonico dei campi (fonte di verità)
    extract.py              # xlsx -> dict per etichetta, derivazione famiglia, normalizzazione tipi
    validate.py             # obbligatori per famiglia + rilevamento valori sospetti
  render/
    template.html           # Jinja2 scheda brandizzata
    styles.css              # CSS paged-media WeasyPrint + brand token
    pdf.py                  # html -> pdf
  assets/
    fonts/                  # .ttf del brand (placeholder Inter finché non arrivano)
    logo.svg
    bolt.svg
    transformers/           # le 4 immagini + anchors.yaml per i callout quote
  .streamlit/config.toml
  requirements.txt
  Dockerfile
  render.yaml
  README.md                 # come girare in locale, deploy, nota sul formato Excel da tenere fisso
```

## Ordine di lavoro suggerito

1. `schema.yaml` completo dalla tabella sopra.
2. `extract.py` + `validate.py` con i 5 file di test reali (li fornisco) come test fixture; verifica famiglia e soppressione righe su ciascuno.
3. `pdf.py` + template + CSS con dati mock, fino a una resa grafica convincente in A4.
4. `app.py` Streamlit che cuce tutto, con il form dinamico e il riepilogo di conferma.
5. Dockerfile, prova locale, deploy Render.

## Cosa NON fare

Niente chiamate a LLM/API per il parsing. Niente database. Niente storage o archivio. Niente email. Niente parsing per coordinata fissa (usa il lookup per etichetta). Niente valori non confermati nel PDF.
