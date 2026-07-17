"""Deterministic Excel parser for the Trafo Elettro datasheet portal.

Reads the fixed key-value extraction sheet by LABEL (not by fixed coordinate),
derives the transformer family, applies the canonical schema, translates
enumerated text values, formats numbers, and selects the configuration image.

No LLM is involved. All logic is deterministic.
"""
from pathlib import Path
import yaml
from openpyxl import load_workbook

from .format import format_value, is_blank, looks_like_coerced_date

SCHEMA_PATH = Path(__file__).with_name("schema.yaml")

SECTION_ORDER = ["general", "electrical", "ratings", "environmental", "dimensions"]
SECTION_TITLE = {
    "general": "General",
    "electrical": "Electrical",
    "ratings": "Ratings",
    "environmental": "Working Conditions",
    "dimensions": "Dimensions & weight",
}


def load_schema(path=SCHEMA_PATH):
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def read_raw(xlsx_path, schema):
    """Return (raw, accessories): key-value fields read by label, and the
    accessory list (rows after the 'Accessori:' marker)."""
    meta = schema["meta"]
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[meta["sheet"]] if meta["sheet"] in wb.sheetnames else wb[wb.sheetnames[0]]
    marker = meta.get("accessories_marker", "Accessori:")
    raw, accessories, in_acc = {}, [], False
    for r in range(meta["first_row"], ws.max_row + 1):
        label = ws.cell(r, meta["label_col"]).value
        value = ws.cell(r, meta["value_col"]).value
        if label is None or not str(label).strip():
            continue
        label = str(label).strip()
        if label == marker:
            in_acc = True
            continue
        if in_acc:
            accessories.append(label)          # accessory descriptions are full sentences
        else:
            raw[label] = value
    return raw, accessories


def derive_family(raw, schema):
    """oil if 'Tipologia olio' present; resin if air cooling; else oil."""
    rules = schema["family_rules"]
    if not is_blank(raw.get("Tipologia olio")):
        return "olio"
    cooling = str(raw.get("Raffreddamento") or "").strip().upper()
    if cooling in [c.upper() for c in rules["air_cooling"]]:
        return "resina"
    if cooling[:1] in rules["oil_prefix"]:
        return "olio"
    serie = str(raw.get("Serie") or "")
    return "resina" if "-R" in serie.upper() else "olio"


IMG_DIR = Path(__file__).resolve().parent.parent / "assets" / "transformers"


def _img_exists(key):
    return bool(key) and (IMG_DIR / f"{key}.png").exists()


def _winding_pair(windings):
    """Coppia di sigle senza numeri (es. 'MV-LV', 'MV-MV', 'HV-MV'); con 3 avvolgimenti
    prende primario + primo secondario."""
    if not windings:
        return "MV-MV"
    if len(windings) == 1:
        b = windings[0]["base"]
        return f"{b}-{b}"
    return f"{windings[0]['base']}-{windings[1]['base']}"


def select_image(raw, family, schema):
    """Compone il nome file dai dati; ricade sui generici / fallback se la combo
    esatta non esiste tra le immagini. Restituisce la chiave (senza .png)."""
    windings = build_windings(raw, schema)
    earthing = is_earthing(raw)
    casa = str(raw.get("Tipo casa") or "").strip().lower()
    cooling = str(raw.get("Tipo sistema raffreddamento") or "").strip().lower()
    oltc = "sottocarico" in str(raw.get("Tipo commutatore") or "").lower()

    cand = []
    if earthing:
        if family == "resina":
            cand.append("Earthing_resina")
        else:
            cassa = "ermetico" if "ermetico" in casa else "conservatore"
            has_bt = any(w["base"] == "LV" for w in windings[1:])  # secondario in bassa tensione
            if has_bt:
                cand.append(f"Earthing_{cassa}_avvolgimento_BT")
            cand.append(f"Earthing_{cassa}")
        cand.append("Earthing_resina")
    elif family == "resina":
        cand.append("Trasformatore_resina")
    else:  # olio
        if "conservatore" in casa and "radiatori" in cooling:
            sig = _winding_pair(windings)
            comm = "OLTC" if oltc else "a_vuoto"
            cand.append(f"Conservatore_radiatori_{sig}_commutatore_{comm}")
            cand.append("Conservatore_radiatori_MV-MV_commutatore_a_vuoto")  # fallback casi speciali
        elif "conservatore" in casa:
            cand.append("Conservatore_onde")
        elif "ermetico" in casa:
            cand.append("Ermetico_onde")
        cand.append("Conservatore_onde")  # fallback generico olio

    for c in cand:
        if _img_exists(c):
            return c
    return cand[0] if cand else None


def translate(label, raw_value, schema):
    """Translate an enumerated italian value to English via value_map.
    Returns (display, translated_ok). translated_ok=False means the value is
    not in the map and must be confirmed in the dynamic form."""
    vmap = schema.get("value_map", {})
    # i campi avvolgimento (Materiale/Tipo avvolg. BT1/BT2/BT3…) condividono lo
    # stesso dizionario dell'MT
    lookup = label
    if label.startswith("Materiale "):
        lookup = "Materiale MT"
    elif label.startswith("Tipo avvolg. "):
        lookup = "Tipo avvolg. MT"
    if lookup in vmap and not is_blank(raw_value):
        key = str(raw_value).strip()
        if key in vmap[lookup]:
            return vmap[lookup][key], True
        toks = key.split()
        if len(toks) > 1 and toks[0] in vmap[lookup]:
            return f"{vmap[lookup][toks[0]]} {' '.join(toks[1:])}", True
        return key, False  # text value present but not mapped -> needs confirmation
    return None, True  # not a mapped field


def build_fields(raw, family, schema):
    """Produce the ordered, sectioned list of display fields for the family."""
    nf = schema["meta"]["number_format"]
    out = []
    for fld in schema["fields"]:
        if not fld.get("include_in_sheet"):
            continue
        if fld["family"] not in ("both", family):
            continue
        if fld["section"] == "header":
            continue
        raw_v = raw.get(fld["it"])
        # text translation (only for mapped fields)
        translated, ok = translate(fld["it"], raw_v, schema)
        display = translated if translated is not None else format_value(raw_v, fld, nf)
        out.append({
            "it": fld["it"], "en": fld["en"], "section": fld["section"],
            "unit": fld.get("unit"), "value": display,
            "blank": is_blank(raw_v), "translated_ok": ok,
        })
    return out


def grouped_sections(fields):
    """Group included, non-blank fields into ordered sections for the template."""
    groups = []
    n = 0
    for sec in SECTION_ORDER:
        rows = [f for f in fields if f["section"] == sec and not f["blank"] and f["value"] not in (None, "")]
        if rows:
            n += 1
            groups.append({"key": sec, "title": SECTION_TITLE[sec], "rows": rows, "idx": n})
    return groups


# --- multi-winding configuration ---
WINDINGS = [
    ("MT",  {"V": "Tensione MT", "V2": "Tensione MT2", "P": "Potenza nominale MT",
             "conn": "Collegamento MT", "ins": "Classe isolamento MT", "mat": "Materiale MT",
             "wt": "Tipo avvolg. MT", "tc": "Classe termica MT", "tr": "Sovratemperatura avvolg. MT"}),
    ("BT1", {"V": "Tensione BT1", "P": "Potenza nominale BT1",
             "conn": "Collegamento BT1", "ins": "Classe isolamento BT1", "mat": "Materiale BT1",
             "wt": "Tipo avvolg. BT1", "tc": "Classe termica BT1", "tr": "Sovratemperatura avvolg. BT1"}),
    ("BT2", {"V": "Tensione BT2", "P": "Potenza nominale BT2",
             "conn": "Collegamento BT2", "ins": "Classe isolamento BT2", "mat": "Materiale BT2",
             "wt": "Tipo avvolg. BT2", "tc": "Classe termica BT2", "tr": "Sovratemperatura avvolg. BT2"}),
    ("BT3", {"V": "Tensione BT3", "P": "Potenza nominale BT3",
             "conn": "Collegamento BT3", "ins": "Classe isolamento BT3", "mat": "Materiale BT3",
             "wt": "Tipo avvolg. BT3", "tc": "Classe termica BT3", "tr": "Sovratemperatura avvolg. BT3"}),
]
WINDING_ROWS = [
    ("Rated power", "P", "kVA", 0),
    ("Voltage", "V", "V", 0),
    ("Connection", "conn", None, None),
    ("Insulation level", "ins", "kV", None),
    ("Winding material", "mat", None, None),
    ("Winding type", "wt", None, None),
    ("Thermal class", "tc", None, None),
    ("Winding temp. rise", "tr", "°C", 0),
]
IMPEDANCES = [
    ("Impedenza di cortocircuito % MT-BT1", "MT", "BT1"),
    ("Impedenza di cortocircuito % MT-BT2", "MT", "BT2"),
    ("Impedenza di cortocircuito % MT-BT3", "MT", "BT3"),
    ("Impedenza di cortocircuito % BT1-BT2", "BT1", "BT2"),
    ("Impedenza di cortocircuito % BT1-BT3", "BT1", "BT3"),
    ("Impedenza di cortocircuito % BT2-BT3", "BT2", "BT3"),
]


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _imp(v, nf):
    if is_blank(v):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return format_value(v, {"decimals": 2}, nf)
    return str(v).strip()  # valori testuali: NA, < 3


def build_windings(raw, schema):
    """Avvolgimenti presenti con sigla LV/MV/HV (numerata se la classe si ripete)."""
    vc = schema["voltage_class"]
    present = []
    for role, m in WINDINGS:
        v = _f(raw.get(m["V"]))
        if v is None or v == 0:
            continue
        ref = v
        v2 = _f(raw.get(m.get("V2")))
        if v2:
            ref = max(v, v2)  # doppia tensione: classifica sulla massima
        base = "LV" if ref <= vc["lv_max"] else ("MV" if ref <= vc["mv_max"] else "HV")
        present.append({"role": role, "base": base, "m": m})
    from collections import Counter
    cnt = Counter(w["base"] for w in present)
    seen = {}
    for w in present:
        b = w["base"]
        if cnt[b] > 1:
            seen[b] = seen.get(b, 0) + 1
            w["label"] = f"{b}{seen[b]}"
        else:
            w["label"] = b
    return present


def _field_by_it(schema, it):
    return next((f for f in schema["fields"] if f["it"] == it), {"it": it})


def get_display(it_label, raw, schema):
    """Single display value for an italian label (translate then format)."""
    fld = _field_by_it(schema, it_label)
    translated, _ = translate(it_label, raw.get(it_label), schema)
    if translated is not None:
        return translated
    return format_value(raw.get(it_label), fld, schema["meta"]["number_format"])


def build_ratings(raw, schema):
    nf = schema["meta"]["number_format"]
    windings = build_windings(raw, schema)
    labels = [w["label"] for w in windings]
    rolelabel = {w["role"]: w["label"] for w in windings}

    rows = []
    for en, key, unit, dec in WINDING_ROWS:
        cells = []
        for w in windings:
            m = w["m"]
            raw_v = raw.get(m[key])
            if dec is not None and isinstance(raw_v, (int, float)) and not isinstance(raw_v, bool):
                val = format_value(raw_v, {"decimals": dec}, nf)
            else:
                val = get_display(m[key], raw, schema)
            if key == "V" and w["role"] == "MT":            # doppia tensione MT: min / max
                v1, v2 = _f(raw.get(m["V"])), _f(raw.get(m.get("V2")))
                if v1 and v2 and v2 != 0:
                    lo, hi = sorted([v1, v2])
                    val = (f"{format_value(lo, {'decimals':0}, nf)} / "
                           f"{format_value(hi, {'decimals':0}, nf)}")
            cells.append(val if val not in (None, "") else "–")
        if any(c != "–" for c in cells):
            rows.append({"label": en, "unit": unit, "cells": cells})

    # tap changer
    taps = []
    tc = get_display("Tipo commutatore", raw, schema)
    if tc:
        taps.append({"label": "Tap changer", "value": tc})
    pp, pm = raw.get("Posizioni + rif. MT"), raw.get("Posizioni - rif. MT")
    if not is_blank(pp) or not is_blank(pm):
        taps.append({"label": "Tap positions (+/-)",
                     "value": f"+{int(pp) if not is_blank(pp) else 0}/-{int(pm) if not is_blank(pm) else 0}"})
    step = raw.get("% gradino rif. MT")
    if not is_blank(step):
        try:
            v = float(step)
            v = v * 100 if v < 1 else v
            taps.append({"label": "Step per tap", "value": f"{format_value(v, {'decimals':2}, nf)} %"})
        except (TypeError, ValueError):
            pass

    return {"windings": labels, "rows": rows, "taps": taps}


def short_circuit_row(raw, schema):
    """Impedenza cc per Electrical, con etichette coppia:
    'MV–LV1: 6% / MV–LV2: 6% / LV1–LV2: 3,5%'."""
    nf = schema["meta"]["number_format"]
    windings = build_windings(raw, schema)
    rolelabel = {w["role"]: w["label"] for w in windings}
    parts = []
    if "BT2" in rolelabel:
        for key, w1, w2 in IMPEDANCES:
            if w1 in rolelabel and w2 in rolelabel:
                v = _imp(raw.get(key), nf)
                if v is not None:
                    parts.append(f"{rolelabel[w1]}–{rolelabel[w2]}: {v}%")
    else:
        v = (_imp(raw.get("Impedenza di cortocircuito % Totale"), nf)
             or _imp(raw.get("Impedenza di cortocircuito % MT-BT1"), nf))
        if v is not None:
            if "MT" in rolelabel and "BT1" in rolelabel:
                parts.append(f"{rolelabel['MT']}–{rolelabel['BT1']}: {v}%")
            else:
                parts.append(f"{v}%")
    if not parts:
        return None
    return {"it": None, "en": "Short-circuit impedance", "section": "electrical", "unit": None,
            "value": " / ".join(parts), "blank": False, "translated_ok": True}


def winding_temp_row(raw, schema, windings):
    """Sovratemperatura avvolgimenti per Working Conditions: valori separati da ' / '."""
    nf = schema["meta"]["number_format"]
    vals = []
    for w in windings:
        v = raw.get(w["m"]["tr"])
        if not is_blank(v):
            vals.append(format_value(v, {"decimals": 0}, nf))
    if not vals:
        return None
    return {"it": None, "en": "Winding temp. rise", "section": "environmental", "unit": "°C",
            "value": " / ".join(vals), "blank": False, "translated_ok": True}


def standards_text(raw):
    xs = [str(raw.get(f"Norma {i} / Regol. {i}")).strip()
          for i in range(1, 5) if not is_blank(raw.get(f"Norma {i} / Regol. {i}"))]
    return " · ".join(xs) if xs else None


def efficiency_row(raw, schema):
    val = raw.get("PEI / MEPS / HEPS in AN/ONAN")
    if is_blank(val):
        return None
    norme = " ".join(str(raw.get(f"Norma {i} / Regol. {i}") or "") for i in range(1, 5)).upper()
    eff = schema["efficiency"]
    if any(t.upper() in norme for t in eff["meps_when"]):
        label = "Efficiency index (MEPS/HEPS)"
    elif any(t.upper() in norme for t in eff["pei_when"]):
        label = "Efficiency index (PEI)"
    else:
        label = "Efficiency index"
    return {"it": None, "en": label, "section": "electrical", "unit": "%",
            "value": format_value(val, {"decimals": 3}, schema["meta"]["number_format"]),
            "blank": False, "translated_ok": True}


def cesi_text(raw, family):
    if family != "resina":
        return None
    cls = raw.get("Classe amb / clim / fuoco")
    if is_blank(cls):
        return None
    return f"{str(cls).replace(' ', '')} type test nr. B4013916"


UM_THRESHOLD = 72.5  # kV, soglia prove aggiuntive


def _um_kv(raw):
    """Um in kV = primo numero della classe isolamento MT ('24 / 50 / 125' -> 24)."""
    v = raw.get("Classe isolamento MT")
    if is_blank(v):
        return None
    first = str(v).split("/")[0].strip().replace(",", ".")
    try:
        return float(first)
    except ValueError:
        return None  # es. cella corrotta in data: soglia non valutabile -> nessuna prova HV


def is_earthing(raw):
    return str(raw.get("Gruppo vettoriale") or "").strip().upper().startswith("Z")


def build_tests(raw, family, earthing_override=None):
    """Prove di routine IEC 60076-1 secondo le casistiche (dal foglio 'Principale').
    Prove interne (SFERA/DFR/insulation resistance/core demag) escluse per scelta."""
    um = _um_kv(raw)
    hv = um is not None and um > UM_THRESHOLD
    oil = family == "olio"
    earth = is_earthing(raw) if earthing_override is None else earthing_override
    oltc = "sottocarico" in str(raw.get("Tipo commutatore") or "").lower()
    note_e = " — only for earthing transformer with secondary winding" if earth else ""

    t = [
        "Measurement of winding resistance",
        "Measurement of voltage ratio and check of phase displacement" + note_e,
        "Measurement of short-circuit impedance and load loss" + note_e,
        "Measurement of no-load loss and current",
        "Dielectric routine tests:",
        "– Applied voltage test (AV)",
        "– " + ("Induced voltage test with PD measurement (IVPD)" if hv
                else "Induced voltage withstand test (IVW)"),
        "– Auxiliary wiring insulation test (AuxW) — only if present",
    ]
    if hv:
        t.append("– Full wave lightning impulse test for the line terminals (LI)")
        t.append("– Line terminal AC withstand voltage test (LTAC) — only for non-uniform insulation")
    if oltc:
        t.append("Tests on on-load tap-changers, where appropriate")
    if oil:
        t.append("Leak testing with pressure for liquid-immersed transformers (tightness test)")
    t.append("Check of the ratio and polarity of built-in current transformers — only if present")
    t.append("Check of core and frame insulation for liquid immersed transformers with core or frame insulation")
    if earth:
        t.append("Measurement of zero-sequence impedance")
    if hv:
        t.append("Additional routine tests (Um > 72.5 kV):")
        t.append("– Determination of capacitances windings-to-earth and between windings")
        t.append("– Measurement of d.c. insulation resistance between each winding to earth and between windings")
        t.append("– Measurement of dissipation factor (tan δ) of the insulation system capacitances")
        t.append("– Measurement of dissolved gases in dielectric liquid from each oil compartment")
        t.append("– Measurement of no-load loss and current at 90 % and 110 % of rated voltage")
    return t


def designation(raw, schema):
    nf = schema["meta"]["number_format"]
    serie = str(raw.get("Serie") or "").strip()
    tail = []
    power = raw.get("Potenza nominale MT")
    if not is_blank(power):
        tail.append(f"{format_value(power, {'decimals':0}, nf)} kVA")
    hv, lv = raw.get("Tensione MT"), raw.get("Tensione BT1")
    if not is_blank(hv) and not is_blank(lv):
        tail.append(f"{format_value(hv, {'decimals':0}, nf)} / {format_value(lv, {'decimals':0}, nf)} V")
    elif not is_blank(hv):
        tail.append(f"{format_value(hv, {'decimals':0}, nf)} V")
    if tail:
        return f"{serie} — " + " · ".join(tail) if serie else " · ".join(tail)
    return serie


def designation_parts(raw, schema):
    """main = potenza; voltage = HV / LV con i valori multipli uniti da '-'
    (es. '10.000-20.000 / 720-400 V'); series = codice serie."""
    nf = schema["meta"]["number_format"]
    serie = str(raw.get("Serie") or "").strip()
    power = raw.get("Potenza nominale MT")
    main = f"{format_value(power, {'decimals':0}, nf)} kVA" if not is_blank(power) else (serie or "Transformer")

    def side(keys):
        vals = []
        for k in keys:
            v = raw.get(k)
            if not is_blank(v) and _f(v) not in (None, 0):
                vals.append(format_value(v, {"decimals": 0}, nf))
        return "-".join(vals)

    hv = side(["Tensione MT", "Tensione MT2"])
    lv = side(["Tensione BT1", "Tensione BT2"])
    if hv and lv:
        voltage = f"{hv} / {lv} V"
    elif hv:
        voltage = f"{hv} V"
    else:
        voltage = ""
    return main, voltage, serie


def parse(xlsx_path, schema=None, overrides=None):
    schema = schema or load_schema()
    raw, accessories = read_raw(xlsx_path, schema)
    if overrides:
        for k, v in overrides.items():
            raw[k] = v  # corrected / supplied values from the dynamic form
    family = derive_family(raw, schema)
    ov = overrides or {}
    image_key = ov.get("__image__") or ov.get("image") or select_image(raw, family, schema)

    fields = build_fields(raw, family, schema)
    windings = build_windings(raw, schema)
    eff = efficiency_row(raw, schema)
    sci = short_circuit_row(raw, schema)
    wtr = winding_temp_row(raw, schema, windings)
    for extra in (eff, sci, wtr):
        if extra:
            fields.append(extra)

    return {
        "raw": raw,
        "family": family,
        "image_key": image_key,
        "designation": designation(raw, schema),
        "fields": fields,
        "sections": grouped_sections(fields),
        "ratings": build_ratings(raw, schema),
        "standards": standards_text(raw),
        "cesi": cesi_text(raw, family),
        "tests": build_tests(raw, family, ov.get("__earthing__")),
        "accessories_excel": accessories,
        "dims": {
            "L": raw.get("Lunghezza trafo"),
            "W": raw.get("Larghezza trafo"),
            "H": raw.get("Altezza trafo"),
        },
    }
