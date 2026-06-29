"""Validation layer. Produces the items the dynamic form must resolve before
the PDF can be generated:
  - missing required fields (for the derived family) + Client + Project
  - suspicious values (Excel date-coercion on insulation-class fields)
  - text values not present in value_map (need translation/confirmation)
  - ambiguous image selection
"""
from .format import is_blank, looks_like_coerced_date

PORTAL_REQUIRED = ["Client", "Project"]  # collected in the portal, not in Excel


def required_fields(schema, family):
    return [f for f in schema["fields"]
            if f.get("required") and f["family"] in ("both", family)]


def validate(parsed, schema, portal_inputs=None):
    portal_inputs = portal_inputs or {}
    raw = parsed["raw"]
    family = parsed["family"]
    items = []  # each: {kind, key, label, unit, raw, options?}

    # 1. missing required from Excel
    for f in required_fields(schema, family):
        if is_blank(raw.get(f["it"])):
            items.append({"kind": "missing", "key": f["it"], "label": f["en"],
                          "unit": f.get("unit"), "raw": None})

    # 2. portal-required (Client / Project)
    for p in PORTAL_REQUIRED:
        if is_blank(portal_inputs.get(p)):
            items.append({"kind": "portal", "key": p, "label": p, "unit": None, "raw": None})

    # 3. suspicious date-coerced values
    for key in schema.get("suspicious", {}).get("date_coerced_fields", []):
        v = raw.get(key)
        if looks_like_coerced_date(v):
            en = next((f["en"] for f in schema["fields"] if f["it"] == key), key)
            items.append({"kind": "suspicious", "key": key, "label": en,
                          "unit": "kV", "raw": str(v)})

    # 4. untranslated enumerated text values
    vmap = schema.get("value_map", {})
    for f in schema["fields"]:
        if not f.get("include_in_sheet") or f["family"] not in ("both", family):
            continue
        label = f["it"]
        if label in vmap and not is_blank(raw.get(label)):
            if str(raw[label]).strip() not in vmap[label]:
                items.append({"kind": "translate", "key": label, "label": f["en"],
                              "unit": None, "raw": str(raw[label]),
                              "options": list(vmap[label].values())})

    # 5. ambiguous image
    if parsed["image_key"] is None:
        items.append({"kind": "image", "key": "image", "label": "Configuration image",
                      "unit": None, "raw": None,
                      "options": ["oil_conservator_radiators", "oil_conservator_corrugated",
                                  "oil_hermetic", "resin_enclosure", "resin_open"]})
    return items
