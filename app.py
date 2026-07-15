"""Trafo Elettro — Technical Datasheet generator (Streamlit, stateless).

Flow: password gate -> upload Excel -> parse -> validate -> dynamic form
(fill / omit) -> summary -> generate PDF -> download.
No database, no storage, no email. No LLM at runtime.
"""
import os
import tempfile
from datetime import date

import streamlit as st

from parser.extract import load_schema, parse
from parser.validate import validate
from render.pdf import render_pdf_modern

OMIT = "__OMIT__"
st.set_page_config(page_title="Trafo Elettro · Datasheet", layout="centered")


# ---------- auth ----------
def gate():
    if st.session_state.get("auth"):
        return True
    st.title("Trafo Elettro · Technical Datasheet")
    pw = st.text_input("Password", type="password")
    if st.button("Enter"):
        if pw and pw == os.environ.get("APP_PASSWORD", ""):
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    return False


def reset():
    for k in ["parsed", "answers", "omitted", "meta", "notes", "uploaded"]:
        st.session_state.pop(k, None)


def main():
    schema = load_schema()
    st.title("Technical Datasheet")
    st.caption("Carica l'Excel di estrazione, completa gli eventuali campi mancanti, genera il PDF.")

    up = st.file_uploader("Excel di estrazione (.xlsx)", type=["xlsx"])
    if up is None:
        reset()
        return

    # persist upload to a temp file (WeasyPrint/openpyxl need a path)
    if st.session_state.get("uploaded") != up.name:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(up.getbuffer()); tmp.flush()
        st.session_state["xlsx_path"] = tmp.name
        st.session_state["uploaded"] = up.name
        st.session_state.pop("answers", None)
        st.session_state.pop("omitted", None)
        st.session_state.pop("acc_text", None)
        st.session_state.pop("tests_text", None)

    # portal meta inputs
    st.subheader("Dati commessa")
    c1, c2, c3 = st.columns(3)
    meta = {
        "client": c1.text_input("Client", st.session_state.get("meta", {}).get("client", "")),
        "project": c2.text_input("Project", st.session_state.get("meta", {}).get("project", "")),
        "date": c3.text_input("Date", st.session_state.get("meta", {}).get("date", date.today().strftime("%d %B %Y"))),
        "offer": c1.text_input("Offer #", ""),
        "pos": c2.text_input("Pos.", ""),
        "product_code": c3.text_input("Product code", ""),
    }

    answers = st.session_state.get("answers", {})
    omitted = set(st.session_state.get("omitted", []))

    # first parse with current answers as overrides
    overrides = {k: v for k, v in answers.items() if v not in (None, "", OMIT)}
    for k in omitted:
        overrides[k] = None
    parsed = parse(st.session_state["xlsx_path"], schema, overrides=overrides)

    portal_inputs = {"Client": meta["client"], "Project": meta["project"]}
    resolved = {k for k, v in answers.items() if v not in (None, "", OMIT)}
    items = [it for it in validate(parsed, schema, portal_inputs)
             if it["key"] not in omitted and it["key"] not in resolved and it["kind"] != "portal"]
    portal_missing = [p for p in ("Client", "Project") if not portal_inputs[p].strip()]

    st.divider()
    st.write(f"**Famiglia rilevata:** {parsed['family']}  ·  **Immagine:** {parsed['image_key'] or '—'}")

    # ---------- dynamic form ----------
    if items or portal_missing:
        st.subheader("Da completare")
        if portal_missing:
            st.warning("Compila Client e Project qui sopra.")
        if st.button("Ometti tutti i campi rimanenti"):
            for it in items:
                if it["kind"] in ("missing", "translate", "suspicious"):
                    omitted.add(it["key"])
            st.session_state["omitted"] = list(omitted)
            st.rerun()

        for it in items:
            st.markdown(f"**{it['label']}**" + (f"  ·  letto: `{it['raw']}`" if it.get("raw") else ""))
            cc1, cc2 = st.columns([3, 1])
            if it.get("options"):
                val = cc1.selectbox("Valore", [""] + it["options"], key=f"v_{it['key']}",
                                    label_visibility="collapsed")
            else:
                unit = f" ({it['unit']})" if it.get("unit") else ""
                val = cc1.text_input(f"Valore{unit}", key=f"v_{it['key']}",
                                     label_visibility="collapsed")
            if cc2.checkbox("Ometti", key=f"o_{it['key']}"):
                omitted.add(it["key"])
            else:
                if val:
                    answers[it["key"]] = val
        st.session_state["answers"] = answers
        st.session_state["omitted"] = list(omitted)
        st.session_state["meta"] = meta
        if st.button("Applica e ricontrolla", type="primary"):
            st.rerun()
        st.info("Completa o ometti i campi, poi premi «Applica e ricontrolla».")
        return

    # ---------- image selector with thumbnail ----------
    tdir = os.path.join(os.path.dirname(__file__), "assets", "transformers")
    OLD = {"oil_conservator_radiators", "oil_conservator_corrugated",
           "oil_hermetic", "resin_enclosure", "resin_open"}
    IMAGES = sorted(f[:-4] for f in os.listdir(tdir)
                    if f.endswith(".png") and f[:-4] not in OLD)
    st.subheader("Immagine trasformatore")
    cur = parsed["image_key"]
    idx = IMAGES.index(cur) if cur in IMAGES else 0
    chosen = st.selectbox("Immagine", IMAGES, index=idx)
    img_path = os.path.join(tdir, f"{chosen}.png")
    if os.path.exists(img_path):
        st.image(img_path, width=260, caption=chosen)
    if chosen != parsed["image_key"]:
        answers["__image__"] = chosen
        st.session_state["answers"] = answers
        parsed["image_key"] = chosen

    # ---------- summary + generate ----------
    st.subheader("Riepilogo valori in scheda")
    for s in parsed["sections"]:
        with st.expander(s["title"], expanded=False):
            for r in s["rows"]:
                st.write(f"- {r['en']}: **{r['value']}**" + (f" {r['unit']}" if r["unit"] else ""))

    # accessories: precompilati dall'Excel, uno per riga, modificabili
    st.subheader("Accessories")
    if "acc_text" not in st.session_state:
        st.session_state["acc_text"] = "\n".join(parsed.get("accessories_excel", []))
    acc_text = st.text_area("Accessori (uno per riga)", height=220, key="acc_text")
    accessories = [l.strip() for l in acc_text.split("\n") if l.strip()]

    # routine tests: derivati dalle regole IEC, precompilati, modificabili
    st.subheader("Routine tests (IEC 60076-1)")
    if "tests_text" not in st.session_state:
        st.session_state["tests_text"] = "\n".join(parsed.get("tests", []))
    tests_text = st.text_area("Prove (una per riga)", height=260, key="tests_text")
    tests = [l.strip() for l in tests_text.split("\n") if l.strip()]

    notes = st.text_area("Notes (facoltative)", st.session_state.get("notes", ""))
    if st.button("Genera PDF", type="primary"):
        out = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        render_pdf_modern(parsed, meta, notes, out.name, accessories=accessories, tests=tests)
        with open(out.name, "rb") as f:
            st.download_button("Scarica la scheda tecnica (PDF)", f.read(),
                               file_name=f"datasheet_{(meta['client'] or 'trafo').replace(' ', '_')}.pdf",
                               mime="application/pdf")


if gate():
    main()
