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
    """Return {italian_label: raw_value} read by label from the key-value sheet."""
    meta = schema["meta"]
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[meta["sheet"]] if meta["sheet"] in wb.sheetnames else wb[wb.sheetnames[0]]
    raw = {}
    for r in range(meta["first_row"], ws.max_row + 1):
        label = ws.cell(r, meta["label_col"]).value
        value = ws.cell(r, meta["value_col"]).value
        if label is not None and str(label).strip():
            raw[str(label).strip()] = value
    return raw


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
        return None  # no Cabina info -> ambiguous, ask in form (enclosure vs open)
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


RATINGS_PAIRS = [
    ("Rated power", "Potenza nominale", "Potenza nominale", "kVA"),
    ("Voltage", "Tensione MT1", "Tensione BT", "V"),
    ("Insulation level", "Classe isolamento MT", "Classe isolamento BT", "kV"),
    ("Winding material", "Materiale MT", "Materiale BT", None),
    ("Winding type", "Tipo avvolg. MT", "Tipo avvolg. BT", None),
    ("Connection", "Collegamento MT", "Collegamento BT", None),
    ("Thermal class", "Classe termica MT", "Classe termica BT", None),
]


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
    pairs = []
    for en, it_hv, it_lv, unit in RATINGS_PAIRS:
        hv = get_display(it_hv, raw, schema)
        lv = get_display(it_lv, raw, schema)
        # doppia tensione MT: se Tensione MT2 valorizzata, mostra MT1 / MT2 sul lato HV
        if it_hv == "Tensione MT1":
            mt2 = raw.get("Tensione MT2")
            if not is_blank(mt2):
                try:
                    if float(mt2) != 0:
                        hv = f"{hv} / {get_display('Tensione MT2', raw, schema)}"
                except (TypeError, ValueError):
                    pass
        if hv is None and lv is None:
            continue
        pairs.append({"label": en, "hv": hv, "lv": lv, "unit": unit})
    taps = []
    tc = get_display("Tipo commutatore", raw, schema)
    if tc:
        taps.append({"label": "Tap changer", "value": tc})
    pp, pm = raw.get("Posizioni + rif. MT1"), raw.get("Posizioni - rif. MT1")
    if not is_blank(pp) or not is_blank(pm):
        taps.append({"label": "Tap positions (+/-)",
                     "value": f"{int(pp) if not is_blank(pp) else '-'} / {int(pm) if not is_blank(pm) else '-'}"})
    step_raw = raw.get("% gradino rif. MT1")
    if not is_blank(step_raw):
        try:
            v = float(step_raw)
            v = v * 100 if v < 1 else v  # <1 = frazione (0.0125 -> 1.25), >=1 = già percentuale
            taps.append({"label": "Step per tap",
                         "value": f"{format_value(v, {'decimals': 2}, nf)} %"})
        except (TypeError, ValueError):
            pass
    return {"pairs": pairs, "taps": taps}


def designation(raw, schema):
    nf = schema["meta"]["number_format"]
    serie = str(raw.get("Serie") or "").strip()
    parts = [serie] if serie else []
    tail = []
    power = raw.get("Potenza nominale")
    if not is_blank(power):
        tail.append(f"{format_value(power, {'decimals':0}, nf)} kVA")
    hv, lv = raw.get("Tensione MT1"), raw.get("Tensione BT")
    if not is_blank(hv) and not is_blank(lv):
        tail.append(f"{format_value(hv, {'decimals':0}, nf)} / {format_value(lv, {'decimals':0}, nf)} V")
    elif not is_blank(hv):
        tail.append(f"{format_value(hv, {'decimals':0}, nf)} V")
    if tail:
        return f"{serie} — " + " · ".join(tail) if serie else " · ".join(tail)
    return serie


def designation_parts(raw, schema):
    """Three separate strings for the modern layout title:
    main = power (big headline),  voltage = HV/LV voltage string,  series = serie code."""
    nf = schema["meta"]["number_format"]
    serie = str(raw.get("Serie") or "").strip()
    power = raw.get("Potenza nominale")
    main = f"{format_value(power, {'decimals':0}, nf)} kVA" if not is_blank(power) else (serie or "Transformer")
    hv, lv = raw.get("Tensione MT1"), raw.get("Tensione BT")
    if not is_blank(hv) and not is_blank(lv):
        voltage = f"{format_value(hv, {'decimals':0}, nf)} / {format_value(lv, {'decimals':0}, nf)} V"
    elif not is_blank(hv):
        voltage = f"{format_value(hv, {'decimals':0}, nf)} V"
    else:
        voltage = ""
    return main, voltage, serie


def parse(xlsx_path, schema=None, overrides=None):
    schema = schema or load_schema()
    raw = read_raw(xlsx_path, schema)
    if overrides:
        for k, v in overrides.items():
            raw[k] = v  # corrected / supplied values from the dynamic form
    family = derive_family(raw, schema)
    ov = overrides or {}
    image_key = ov.get("__image__") or ov.get("image") or select_image(raw, family, schema)
    fields = build_fields(raw, family, schema)
    return {
        "raw": raw,
        "family": family,
        "image_key": image_key,
        "designation": designation(raw, schema),
        "fields": fields,
        "sections": grouped_sections(fields),
        "ratings": build_ratings(raw, schema),
        "dims": {
            "L": raw.get("Lunghezza trafo"),
            "W": raw.get("Larghezza trafo"),
            "H": raw.get("Altezza trafo"),
        },
    }
