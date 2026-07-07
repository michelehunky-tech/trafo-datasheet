"""Calibra le quote assonometriche di TUTTE le immagini usando la geometria
di oil_conservator_radiators, adattata alla dimensione reale di ciascun PNG.
Uso: python3 calibrate_all.py  (dalla root del repo)
"""
import struct
import os
import yaml

# calibrazione di riferimento, misurata su oil_conservator_radiators (946x759)
REF_W, REF_H = 946, 759
VB_TOP, VB_BOTTOM = 20, 66
REF = {
    "L": {"start": [96, 600], "dir": [0.866, 0.5], "length": 455, "rot": 30},
    "W": {"start": [556, 812], "dir": [0.866, -0.5], "length": 395, "rot": -30},
    "H": {"start": [52, 212], "dir": [0.0, 1.0], "length": 258, "rot": -90},
}
DIR = "assets/transformers"
NAMES = ["oil_conservator_radiators", "oil_conservator_corrugated",
         "oil_hermetic", "resin_enclosure", "resin_open"]


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    return struct.unpack(">II", head[16:24])  # width, height dall'header PNG


def scaled(axis, s):
    a = REF[axis]
    return {"start": [round(a["start"][0] * s), round(a["start"][1] * s)],
            "dir": a["dir"], "length": round(a["length"] * s), "rot": a["rot"]}


out = {}
for n in NAMES:
    p = os.path.join(DIR, f"{n}.png")
    if not os.path.exists(p):
        print("manca:", p)
        continue
    w, h = png_size(p)
    s = w / REF_W  # scala uniforme (assume stesso aspect ratio del riferimento)
    out[n] = {
        "image": {"w": w, "h": h},
        "viewbox": f"0 {round(VB_TOP * s)} {w} {round(h + VB_BOTTOM * s)}",
        "L": scaled("L", s), "W": scaled("W", s), "H": scaled("H", s),
    }
    print(f"{n}: {w}x{h}  scala={s:.3f}")

with open("render/anchors.yaml", "w") as f:
    f.write("# calibrazione assonometrica — tutte le immagini usano la geometria di oil_conservator_radiators\n")
    yaml.safe_dump(out, f, sort_keys=False, allow_unicode=True)
print(f"\nrender/anchors.yaml scritto per {len(out)} immagini.")
