"""Trafo Elettro — Technical Datasheet generator (Streamlit, stateless).

Flow: password gate -> upload Excel -> parse -> validate -> dynamic form
(fill / omit) -> summary -> generate PDF -> download.
No database, no storage, no email. No LLM at runtime.
"""
import os
import base64
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

from parser.extract import load_schema, parse
from parser.validate import validate
from render.pdf import render_pdf_modern

OMIT = "__OMIT__"
ASSETS = Path(__file__).with_name("assets")
st.set_page_config(page_title="Trafo Elettro · Datasheet",
                   page_icon=str(ASSETS / "logo.png") if (ASSETS / "logo.png").exists() else None,
                   layout="centered")


def _b64(path):
    try:
        return base64.b64encode(Path(path).read_bytes()).decode("ascii")
    except Exception:
        return ""


DARK_VARS = """
:root{
  --bg:#0E0F13; --panel:#171921; --ink:#F1F2F5; --ink-soft:#9AA0AA;
  --line:#2A2D36; --accent:#FFD000; --accent-ink:#14151A;
  --stripe:rgba(255,255,255,0.02);
  --shadow:0 10px 40px -12px rgba(0,0,0,0.6), 0 2px 8px -2px rgba(0,0,0,0.3);
}
"""


def inject_style():
    """Inietta il CSS globale. Il tema attivo sovrascrive le variabili su :root
    (niente JavaScript: Streamlit non lo eseguirebbe)."""
    theme = st.session_state.get("theme", "light")
    css = (ASSETS / "app_style.css").read_text() if (ASSETS / "app_style.css").exists() else ""
    override = DARK_VARS if theme == "dark" else ""
    st.markdown(f"<style>{css}\n{override}</style>", unsafe_allow_html=True)


def logo_img_tag(cls="logo", height=None):
    b = _b64(ASSETS / "logo.png")
    if not b:
        return ""
    style = f' style="height:{height}"' if height else ""
    return f'<img class="{cls}" src="data:image/png;base64,{b}"{style}>'


def reset_all():
    """Cancella tutto tranne autenticazione e tema, per iniziare una nuova scheda."""
    keep = {"auth", "theme"}
    for k in list(st.session_state.keys()):
        if k not in keep:
            del st.session_state[k]


@st.dialog("Scheda tecnica pronta")
def pdf_dialog():
    data, fname = st.session_state.get("pdf_ready", (None, None))
    if not data:
        return
    st.markdown('<div class="te-dialog-ok">✓</div>'
                '<p style="text-align:center;margin:0 0 18px;color:var(--ink-soft);">'
                'Il PDF è stato generato correttamente.</p>',
                unsafe_allow_html=True)
    st.download_button("⬇  Scarica il PDF", data, file_name=fname,
                       mime="application/pdf", use_container_width=True)
    if st.button("Chiudi", use_container_width=True, key="__close_dialog"):
        st.session_state.pop("pdf_ready", None)
        st.rerun()


def header():
    """Logo a sinistra; a destra: Nuova scheda · icona tema · logout."""
    current = st.session_state.get("theme", "light")
    c_logo, c_new, c_theme, c_out = st.columns([6, 1.6, 0.7, 0.7])
    with c_logo:
        st.markdown(f'<div class="te-brand">{logo_img_tag(height="34px")}</div>',
                    unsafe_allow_html=True)
    with c_new:
        if st.button("＋ Nuova scheda", key="__new", use_container_width=True):
            reset_all()
            st.rerun()
    with c_theme:
        st.markdown('<div class="te-iconbtn">', unsafe_allow_html=True)
        icon = "🌙" if current == "light" else "☀️"
        if st.button(icon, key="__theme_btn", help="Cambia tema chiaro/scuro",
                     use_container_width=True):
            st.session_state["theme"] = "dark" if current == "light" else "light"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with c_out:
        st.markdown('<div class="te-iconbtn">', unsafe_allow_html=True)
        if st.button("⎋", key="__logout", help="Esci",
                     use_container_width=True):
            reset_all()
            st.session_state["auth"] = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<hr style="margin:2px 0 18px;border:none;border-top:1px solid var(--line);">',
                unsafe_allow_html=True)


# ---------- auth ----------
def gate():
    inject_style()
    if st.session_state.get("auth"):
        return True
    st.markdown(
        f'<div class="te-gate"><div class="te-card">'
        f'{logo_img_tag("logo")}'
        f'<h1>Technical Datasheet</h1>'
        f'<p class="sub">Enter password to continue</p>'
        f'</div></div>',
        unsafe_allow_html=True,
    )
    # widget dentro il flusso normale, ma stilati dal CSS del gate
    with st.container():
        st.markdown('<div class="te-gate-form">', unsafe_allow_html=True)
        pw = st.text_input("Password", type="password", label_visibility="collapsed",
                           placeholder="Password")
        if st.button("Enter", use_container_width=True):
            if pw and pw == os.environ.get("APP_PASSWORD", ""):
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("Wrong password.")
        st.markdown('</div>', unsafe_allow_html=True)
    return False


def reset():
    for k in ["parsed", "answers", "omitted", "meta", "notes", "uploaded"]:
        st.session_state.pop(k, None)


def main():
    inject_style()
    header()
    schema = load_schema()
    st.markdown('<h1 style="margin:0 0 6px;">Technical Datasheet</h1>'
                '<p style="color:var(--ink-soft);margin:0 0 22px;">'
                'Carica l\'Excel di estrazione, completa gli eventuali campi mancanti, genera il PDF.</p>',
                unsafe_allow_html=True)

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
        "revision": c2.text_input("Revision #", ""),
        "pos": c3.text_input("Pos.", ""),
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
    with st.expander("Ratings", expanded=True):
        rt = parsed["ratings"]
        cols = "  |  ".join(rt["windings"])
        st.write(f"**Avvolgimenti:** {cols}")
        for row in rt["rows"]:
            st.write(f"- {row['label']}: **{' | '.join(row['cells'])}**" + (f" {row['unit']}" if row["unit"] else ""))
        for t in rt["taps"]:
            st.write(f"- {t['label']}: **{t['value']}**")
    for s in parsed["sections"]:
        with st.expander(s["title"], expanded=False):
            for r in s["rows"]:
                st.write(f"- {r['en']}: **{r['value']}**" + (f" {r['unit']}" if r["unit"] else ""))

    # accessories/test: value= precompila (mostra subito le righe); key legata al file
    # + nonce per il ripristino.
    uploaded_id = str(st.session_state.get("uploaded", "file")).replace(".", "_").replace(" ", "_")
    excel_acc = "\n".join(parsed.get("accessories_excel", []))
    excel_tests = "\n".join(parsed.get("tests", []))

    st.subheader("Accessories")
    ca1, ca2 = st.columns([5, 1])
    ca1.caption(f"{len(parsed.get('accessories_excel', []))} letti dall'Excel · uno per riga, modificabili")
    if ca2.button("↺ Excel", key="__reset_acc", help="Ripristina dagli accessori dell'Excel"):
        st.session_state["acc_nonce"] = st.session_state.get("acc_nonce", 0) + 1
        st.rerun()
    acc_text = st.text_area("Accessori", value=excel_acc, height=220,
                            key=f"accw_{uploaded_id}_{st.session_state.get('acc_nonce', 0)}",
                            label_visibility="collapsed")
    accessories = [l.strip() for l in acc_text.split("\n") if l.strip()]

    st.subheader("Routine tests (IEC 60076-1)")
    ct1, ct2 = st.columns([5, 1])
    ct1.caption(f"{len(parsed.get('tests', []))} prove derivate dalle regole IEC · una per riga, modificabili")
    if ct2.button("↺ Regole", key="__reset_tests", help="Ripristina le prove derivate"):
        st.session_state["tests_nonce"] = st.session_state.get("tests_nonce", 0) + 1
        st.rerun()
    tests_text = st.text_area("Prove", value=excel_tests, height=260,
                              key=f"testsw_{uploaded_id}_{st.session_state.get('tests_nonce', 0)}",
                              label_visibility="collapsed")
    tests = [l.strip() for l in tests_text.split("\n") if l.strip()]

    notes = st.text_area("Notes (facoltative)", st.session_state.get("notes", ""))
    if st.button("Genera PDF", type="primary"):
        with st.spinner("Generazione della scheda tecnica in corso…"):
            out = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            render_pdf_modern(parsed, meta, notes, out.name, accessories=accessories, tests=tests)
            with open(out.name, "rb") as f:
                pdf_bytes = f.read()
        fname = f"datasheet_{(meta['client'] or 'trafo').replace(' ', '_')}.pdf"
        st.session_state["pdf_ready"] = (pdf_bytes, fname)
        pdf_dialog()


if gate():
    main()
