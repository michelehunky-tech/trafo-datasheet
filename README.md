# Trafo Elettro — Generatore schede tecniche

Web app interna e condivisa: carichi l'Excel di estrazione dell'ufficio tecnico,
l'app lo interpreta in modo deterministico, valida i campi obbligatori, chiede in un
form dinamico solo ciò che manca (con opzione «Ometti»), e genera un PDF di scheda
tecnica brandizzato Trafo Elettro. Niente database, niente storage, niente email,
nessun LLM a runtime.

## Stato

Verificato e funzionante end-to-end sui file di test reali:
- parser deterministico per etichetta (`parser/extract.py`)
- schema canonico dei 72 campi con mapping IT→EN, famiglie, obbligatori condizionali, `value_map`, regole immagine (`parser/schema.yaml`)
- validazione: obbligatori per famiglia, valori sospetti (date), valori da tradurre, immagine ambigua (`parser/validate.py`)
- renderer HTML/CSS → PDF con WeasyPrint, stile brand verificato, quote assonometriche sull'immagine (`render/`)
- app Streamlit con auth, form dinamico, riepilogo, download (`app.py`)
- font Zalando Sans + SemiExpanded già inclusi in `assets/fonts/`
- logo e Image 1 (oil_conservator_radiators) già inclusi e calibrati

Da completare (lavoro residuo, tutto già predisposto):
1. Aggiungere le altre 3 immagini in `assets/transformers/` con i nomi indicati e calibrarle in `render/anchors.yaml` (Image 1 è il modello già fatto).
2. Concordare con l'ufficio tecnico il formato Excel fisso: celle `Classe isolamento MT/BT` formattate come **Testo** (evita la conversione in data) e vocabolario chiuso per i campi con `value_map`.
3. Decidere quali dei campi marcati `_internal` nello schema vanno eventualmente mostrati (basta `include_in_sheet: true`).

## Avvio in locale

```bash
pip install -r requirements.txt
# WeasyPrint richiede librerie di sistema: vedi Dockerfile (Pango, Cairo, gdk-pixbuf).
export APP_PASSWORD="scegli-una-password"
streamlit run app.py
```

## Deploy su Render

Servizio Docker (vedi `Dockerfile` e `render.yaml`). Imposta la variabile d'ambiente
`APP_PASSWORD` nel dashboard. App stateless: nessun persistent disk. Health check su `/`.

## Come funziona (pipeline)

```
upload .xlsx
  -> read_raw (lookup per etichetta su Foglio1, colonne A/B)
  -> derive_family (olio/resina da Tipologia olio / Raffreddamento)
  -> select_image (Tipo casa + Tipo sistema raffreddamento)
  -> build_fields + build_ratings (mapping EN, traduzione value_map, formato numerico)
  -> validate (mancanti per famiglia + Client/Project + sospetti + da tradurre)
  -> form dinamico (compila / ometti)  -> riepilogo
  -> render_pdf (Jinja2 + WeasyPrint, quote assonometriche)
  -> download
```

## Regole chiave (sono dati, non codice: si cambiano in `schema.yaml`)

- **Famiglia**: olio se `Tipologia olio` valorizzato; resina se `Raffreddamento` ad aria (AN/AF); altrimenti olio.
- **Immagine**: resina→`resin_open`; olio+Ermetico→`oil_hermetic`; olio+Conservatore+Radiatori→`oil_conservator_radiators`; olio+Conservatore+Onde→`oil_conservator_corrugated`. Se ambigua, la chiede il form.
- **value_map**: traduce i valori enumerati IT→EN (Conservatore→Conservator, Reg. sottocarico→OLTC, Strati→Layer, …). Valore non in mappa → il form lo fa confermare.
- **Date-bug**: `Classe isolamento MT/BT` letti come data → segnalati come sospetti nel form.
- **Formato numerico**: inglese (punto decimale, virgola migliaia), zeri finali rimossi, scala per campo (es. `% gradino` ×100).

## Struttura

```
app.py                  Streamlit: auth, upload, form dinamico, riepilogo, genera
parser/schema.yaml      fonte di verità: campi, mapping, famiglie, value_map, regole immagine
parser/extract.py       parsing deterministico, famiglia, traduzione, ratings HV/LV
parser/validate.py      obbligatori per famiglia + sospetti + da tradurre + immagine
parser/format.py        formattazione numerica (decimali, migliaia, scala)
render/template.html    template Jinja2 della scheda
render/styles.css       stile brand verificato (token, @font-face, tabelle, header/footer, tick diagonale)
render/anchors.yaml     calibrazione assonometrica per immagine (quote L/W/H)
render/pdf.py           costruisce l'overlay quote e renderizza in PDF
assets/fonts/           Zalando Sans + SemiExpanded (TTF, già inclusi)
assets/logo.svg         logo Trafo Elettro
assets/transformers/    le 4 immagini (Image 1 inclusa; le altre 3 da aggiungere)
gen_schema.py           generatore opzionale dello schema (tool di sviluppo)
SPEC.md                 specifica completa di progetto (rationale e regole)
Dockerfile, render.yaml, requirements.txt, .streamlit/config.toml
```

## Note WeasyPrint

Font self-hostati via `@font-face` (no Google Fonts a runtime). Diagonali con SVG o
gradiente, non `clip-path`. Verificare sempre il rendering nel container Docker, non
solo in locale, perché font e librerie di sistema devono coincidere.
