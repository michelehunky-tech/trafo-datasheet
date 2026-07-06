"""Render the datasheet HTML to PDF with WeasyPrint, including the
axonometric dimension overlay on the configuration image."""
from pathlib import Path
import yaml
from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML

ROOT = Path(__file__).resolve().parent.parent
RENDER_DIR = Path(__file__).resolve().parent
ANCHORS = yaml.safe_load(open(RENDER_DIR / "anchors.yaml", encoding="utf-8"))


def _fmt(v):
    try:
        return f"{int(float(v)):,}".replace(",", ".")   # European thousands: dot
    except (TypeError, ValueError):
        return str(v)


def build_figure_svg(image_key, dims):
    """Return inline SVG: image + 3 axis-parallel dimension lines (no arrows)
    with measures. Falls back to image + caption when no calibration exists."""
    rel = f"assets/transformers/{image_key}.png"
    cal = ANCHORS.get(image_key)
    if not cal:
        return (f'<img class="figimg" src="{rel}">'
                f'<div class="dimcap"><span class="lbl">Overall L × W × H</span>'
                f'{_fmt(dims["L"])} × {_fmt(dims["W"])} × {_fmt(dims["H"])} mm</div>')

    w, h = cal["image"]["w"], cal["image"]["h"]
    lines, labels = [], []
    for axis, value in (("L", dims["L"]), ("W", dims["W"]), ("H", dims["H"])):
        a = cal[axis]
        sx, sy = a["start"]
        ex = sx + a["dir"][0] * a["length"]
        ey = sy + a["dir"][1] * a["length"]
        mx, my = (sx + ex) / 2, (sy + ey) / 2
        lines.append(f'<line x1="{sx:.0f}" y1="{sy:.0f}" x2="{ex:.0f}" y2="{ey:.0f}"/>')
        labels.append(
            f'<g transform="translate({mx:.0f},{my:.0f}) rotate({a["rot"]})">'
            f'<rect x="-64" y="-16" width="128" height="29" rx="2" fill="#fff"/>'
            f'<text x="0" y="5" text-anchor="middle">{axis} {_fmt(value)} mm</text></g>')
    return (
        f'<svg class="figsvg" viewBox="{cal["viewbox"]}" xmlns="http://www.w3.org/2000/svg" '
        f'xmlns:xlink="http://www.w3.org/1999/xlink">'
        f'<image href="{rel}" x="0" y="0" width="{w}" height="{h}"/>'
        f'<g stroke="#1D1E1B" stroke-width="2.4" stroke-linecap="round" fill="none">{"".join(lines)}</g>'
        f'<g font-family="Zalando Sans" font-weight="500" font-size="25" fill="#1D1E1B">{"".join(labels)}</g>'
        f'</svg>')


def render_pdf(parsed, meta, notes, out_path, accessories=None):
    env = Environment(loader=FileSystemLoader(str(RENDER_DIR)),
                      autoescape=select_autoescape(["html"]))
    tpl = env.get_template("template.html")
    figure_svg = build_figure_svg(parsed["image_key"], parsed["dims"]) if parsed["image_key"] else ""
    html = tpl.render(
        designation=parsed["designation"],
        meta=meta,
        sections={s["key"]: s for s in parsed["sections"]},
        ratings=parsed["ratings"],
        figure_svg=figure_svg,
        accessories=[a for a in (accessories or []) if (a.get("name") or "").strip()],
        notes=notes,
    )
    HTML(string=html, base_url=str(ROOT)).write_pdf(out_path)
    return out_path


def render_pdf_modern(parsed, meta, notes, out_path, accessories=None, certifications=None):
    """Alternative layout: Space Grotesk, centered hero figure, numbered sections."""
    from parser.extract import designation_parts, load_schema
    env = Environment(loader=FileSystemLoader(str(RENDER_DIR)),
                      autoescape=select_autoescape(["html"]))
    tpl = env.get_template("template_modern.html")
    figure_svg = build_figure_svg(parsed["image_key"], parsed["dims"]) if parsed["image_key"] else ""
    main, voltage, serie = designation_parts(parsed["raw"], load_schema())
    html = tpl.render(
        designation_main=main,
        designation_voltage=voltage,
        designation_series=serie,
        meta=meta,
        sections={s["key"]: s for s in parsed["sections"]},
        ratings=parsed["ratings"],
        figure_svg=figure_svg,
        accessories=[a for a in (accessories or []) if (a.get("name") or "").strip()],
        certifications=[c for c in (certifications or []) if (c.get("name") or "").strip()],
        notes=notes,
    )
    HTML(string=html, base_url=str(ROOT)).write_pdf(out_path)
    return out_path


def render_pdf_industrial(parsed, meta, notes, out_path, accessories=None):
    """Alternative layout: IBM Plex Sans/Mono, drawing-cartouche register, industrial."""
    from parser.extract import designation_parts, load_schema
    env = Environment(loader=FileSystemLoader(str(RENDER_DIR)),
                      autoescape=select_autoescape(["html"]))
    tpl = env.get_template("template_industrial.html")
    figure_svg = build_figure_svg(parsed["image_key"], parsed["dims"]) if parsed["image_key"] else ""
    main, voltage, serie = designation_parts(parsed["raw"], load_schema())
    sub = "  ·  ".join(p for p in [serie, voltage] if p)
    html = tpl.render(
        designation_main=main,
        designation_sub=sub,
        meta=meta,
        sections={s["key"]: s for s in parsed["sections"]},
        ratings=parsed["ratings"],
        figure_svg=figure_svg,
        accessories=[a for a in (accessories or []) if (a.get("name") or "").strip()],
        notes=notes,
    )
    HTML(string=html, base_url=str(ROOT)).write_pdf(out_path)
    return out_path
