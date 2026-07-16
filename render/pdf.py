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
    """Immagine + misure L×W×H con label e icona cartesiano. Nessuna quota sul disegno."""
    rel = f"assets/transformers/{image_key}.png"
    return (f'<img class="figimg" src="{rel}">'
            f'<div class="dimcap">'
            f'<div class="lbl">Overall L × W × H</div>'
            f'<div class="meas"><img class="axesicon" src="assets/cartesiano.png">'
            f'<span>{_fmt(dims["L"])} × {_fmt(dims["W"])} × {_fmt(dims["H"])} mm</span></div>'
            f'</div>')


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


def render_pdf_modern(parsed, meta, notes, out_path, accessories=None, tests=None):
    """Alternative layout: Space Grotesk, centered hero figure, numbered sections."""
    from parser.extract import designation_parts, load_schema
    env = Environment(loader=FileSystemLoader(str(RENDER_DIR)),
                      autoescape=select_autoescape(["html"]))
    tpl = env.get_template("template_modern.html")
    figure_svg = build_figure_svg(parsed["image_key"], parsed["dims"]) if parsed["image_key"] else ""
    main, voltage, serie = designation_parts(parsed["raw"], load_schema())
    acc = accessories if accessories is not None else parsed.get("accessories_excel", [])
    acc = [a for a in acc if str(a).strip()]
    tst = tests if tests is not None else parsed.get("tests", [])
    tst = [t for t in tst if str(t).strip()]
    html = tpl.render(
        designation_main=main,
        designation_voltage=voltage,
        designation_series=serie,
        meta=meta,
        sections={s["key"]: s for s in parsed["sections"]},
        ratings=parsed["ratings"],
        figure_svg=figure_svg,
        cesi=parsed.get("cesi"),
        standards=parsed.get("standards"),
        accessories=acc,
        tests=tst,
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
