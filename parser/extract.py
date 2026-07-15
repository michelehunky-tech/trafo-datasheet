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


def select_image(raw, family, schema):
    """Return image key from rules, or None if ambiguous (then ask in form)."""
    rules = schema["image_rules"]
    if family == "resina":
        cab = raw.get("Cabina")
        if not is_blank(cab):
            truthy = str(cab).strip().lower() in ("si", "sì", "yes", "y", "true", "1", "x")
            return rules["resina"]["cabina_true" if truthy else "cabina_false"]
        return rules["resina"]["cabina_false"]  # default resin_open, sovrascrivibile nel form
    casa = str(raw.get("Tipo casa") or "").strip()
    branch = rules["olio"].get(casa)
    if isinstance(branch, str):
        return branch
    if isinstance(branch, dict):
        cooling_sys = str(raw.get("Tipo sistema raffreddamento") or "").strip()
        return branch.get(cooling_sys)  # may be None -> ambiguous
    return None


def translate(label, raw_value, schema):
    """Translate an enumerated italian value to English via value_map.
    Returns (display, translated_ok). translated_ok=False means the value is
    not in the map and must be confirmed in the dynamic form."""
    vmap = schema.get("value_map", {})
    if label in vmap and not is_blank(raw_value):
        key = str(raw_value).strip()
        if key in vmap[label]:
            return vmap[label][key], True
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
]
WINDING_ROWS = [
    ("Rated power", "P", "kVA"),
    ("Voltage", "V", "V"),
    ("Connection", "conn", None),
    ("Insulation level", "ins", "kV"),
    ("Winding material", "mat", None),
    ("Winding type", "wt", None),
    ("Thermal class", "tc", None),
    ("Winding temp. rise", "tr", "°C"),
]
IMPEDANCES = [
    ("Impedenza di cortocircuito % MT-BT1", "MT", "BT1"),
    ("Impedenza di cortocircuito % MT-BT2", "MT", "BT2"),
    ("Impedenza di cortocircuito % BT1-BT2", "BT1", "BT2"),
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
    for en, key, unit in WINDING_ROWS:
        cells = []
        for w in windings:
            m = w["m"]
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

    # impedenze: coppie se due secondari, altrimenti valore singolo
    impedances = []
    if "BT2" in rolelabel:
        for key, w1, w2 in IMPEDANCES:
            if w1 not in rolelabel or w2 not in rolelabel:
                continue
            val = _imp(raw.get(key), nf)
            if val is not None:
                impedances.append({"label": f"Short-circuit impedance ({rolelabel[w1]}–{rolelabel[w2]})",
                                   "value": val, "unit": "%"})
    else:
        val = (_imp(raw.get("Impedenza di cortocircuito % Totale"), nf)
               or _imp(raw.get("Impedenza di cortocircuito % MT-BT1"), nf))
        if val is not None:
            impedances.append({"label": "Short-circuit impedance", "value": val, "unit": "%"})

    # tap changer
    taps = []
    tc = get_display("Tipo commutatore", raw, schema)
    if tc:
        taps.append({"label": "Tap changer", "value": tc})
    pp, pm = raw.get("Posizioni + rif. MT"), raw.get("Posizioni - rif. MT")
    if not is_blank(pp) or not is_blank(pm):
        taps.append({"label": "Tap positions (+/-)",
                     "value": f"{int(pp) if not is_blank(pp) else '-'} / {int(pm) if not is_blank(pm) else '-'}"})
    step = raw.get("% gradino rif. MT")
    if not is_blank(step):
        try:
            v = float(step)
            v = v * 100 if v < 1 else v
            taps.append({"label": "Step per tap", "value": f"{format_value(v, {'decimals':2}, nf)} %"})
        except (TypeError, ValueError):
            pass

    return {"windings": labels, "rows": rows, "impedances": impedances, "taps": taps}


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
            "value": format_value(val, {"decimals": 2}, schema["meta"]["number_format"]),
            "blank": False, "translated_ok": True}


def standards_row(raw):
    xs = [str(raw.get(f"Norma {i} / Regol. {i}")).strip()
          for i in range(1, 5) if not is_blank(raw.get(f"Norma {i} / Regol. {i}"))]
    if not xs:
        return None
    return {"it": None, "en": "Standards / Regulations", "section": "general", "unit": None,
            "value": " · ".join(xs), "blank": False, "translated_ok": True}


def cesi_text(raw, family):
    if family != "resina":
        return None
    cls = raw.get("Classe amb / clim / fuoco")
    if is_blank(cls):
        return None
    return f"{str(cls).replace(' ', '')} type test nr. B4013916"


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
    """main = potenza (titolone), voltage = MT/BT, series = codice serie."""
    nf = schema["meta"]["number_format"]
    serie = str(raw.get("Serie") or "").strip()
    power = raw.get("Potenza nominale MT")
    main = f"{format_value(power, {'decimals':0}, nf)} kVA" if not is_blank(power) else (serie or "Transformer")
    hv, lv = raw.get("Tensione MT"), raw.get("Tensione BT1")
    if not is_blank(hv) and not is_blank(lv):
        voltage = f"{format_value(hv, {'decimals':0}, nf)} / {format_value(lv, {'decimals':0}, nf)} V"
    elif not is_blank(hv):
        voltage = f"{format_value(hv, {'decimals':0}, nf)} V"
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
    std = standards_row(raw)
    eff = efficiency_row(raw, schema)
    if std:
        fields.append(std)
    if eff:
        fields.append(eff)

    return {
        "raw": raw,
        "family": family,
        "image_key": image_key,
        "designation": designation(raw, schema),
        "fields": fields,
        "sections": grouped_sections(fields),
        "ratings": build_ratings(raw, schema),
        "cesi": cesi_text(raw, family),
        "accessories_excel": accessories,
        "dims": {
            "L": raw.get("Lunghezza trafo"),
            "W": raw.get("Larghezza trafo"),
            "H": raw.get("Altezza trafo"),
        },
    }
