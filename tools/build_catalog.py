#!/usr/bin/env python3
"""Генерирует превью каталога: assets/catalog/<key>.jpg с номерами фигур (#id)
и reference/catalog.md. Запускать после tools/build_templates.py.
Нужны LibreOffice (soffice) и pdftoppm."""
import glob
import os
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont
from pptx.util import Emu

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(ROOT, "skills", "gsom-presentation")
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from catalog import CATALOG, DESCRIPTIONS  # noqa: E402
from gsom_deck import TEMPLATE, Deck, abs_box, ensure_fonts, fillable, shape_kind, slide_key  # noqa: E402

WIDTH = 1100
OUT = os.path.join(SKILL, "assets", "catalog")
FONT_PATH = os.path.join(SKILL, "assets", "fonts", "Inter-Variable.ttf")


def main():
    ensure_fonts()
    os.makedirs(OUT, exist_ok=True)
    for old in glob.glob(os.path.join(OUT, "*.jpg")):
        os.remove(old)
    deck = Deck()
    slides = list(deck.prs.slides)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, TEMPLATE],
                       check=True, capture_output=True)
        pdf = glob.glob(os.path.join(tmp, "*.pdf"))[0]
        subprocess.run(["pdftoppm", "-scale-to-x", str(WIDTH), "-scale-to-y", "-1", "-png", pdf,
                        os.path.join(tmp, "s")], check=True)
        pngs = sorted(glob.glob(os.path.join(tmp, "s-*.png")))
        assert len(pngs) == len(slides), (len(pngs), len(slides))
        font = ImageFont.truetype(FONT_PATH, 12)
        font.set_variation_by_axes([14, 700]) if hasattr(font, "set_variation_by_axes") else None
        emu_w = deck.prs.slide_width
        for slide, png in zip(slides, pngs):
            key = slide_key(slide)
            im = Image.open(png).convert("RGB")
            k = im.width / emu_w
            d = ImageDraw.Draw(im, "RGBA")
            placed = []
            lib = key.startswith("lib-")
            for sh, depth in fillable(slide):
                kind = shape_kind(sh)
                if kind == "group" and not key.startswith("lib-"):
                    continue          # группы подписываем только в библиотеках (иконки)
                box = abs_box(sh)
                if not box:
                    continue
                x, y, w, h = (int(v * k) for v in box)
                if not lib:
                    d.rectangle([x, y, x + w, y + h], outline=(0, 120, 255, 150), width=1)
                label = f"#{sh.shape_id}"
                tw = int(d.textlength(label, font=font)) + 6
                lx, ly = (max(0, x + (w - tw) // 2), y + h + 1) if lib else (max(0, x), max(0, y))
                while any(abs(lx - px) < tw and abs(ly - py) < 16 for px, py in placed):
                    ly += 16
                placed.append((lx, ly))
                d.rectangle([lx, ly, lx + tw, ly + 16], fill=(0, 90, 220, 230))
                d.text((lx + 3, ly + 1), label, fill="white", font=font)
            im.save(os.path.join(OUT, f"{key}.jpg"), quality=72, optimize=True)
    write_md()
    print(f"{len(slides)} превью -> {OUT}")


def write_md():
    lines = [
        "# Каталог слайдов шаблона ВШМ СПбГУ",
        "",
        "Сгенерировано `tools/build_catalog.py`. Превью с номерами фигур — `assets/catalog/<key>.jpg`",
        "(синие метки `#id` — то, что можно заполнить через `fields`). Точный список фигур:",
        "`python3 scripts/gsom_deck.py fields <key>`.",
        "",
    ]
    section = None
    for key in CATALOG.values():
        sec, desc = DESCRIPTIONS.get(key, ("?", ""))
        if sec != section:
            lines += ["", f"## {sec}", "", "| Ключ | Назначение | Превью |", "|---|---|---|"]
            section = sec
        lines.append(f"| `{key}` | {desc} | [jpg](../assets/catalog/{key}.jpg) |")
    lines += ["", "## Служебные", "", "| Ключ | Назначение |", "|---|---|",
              "| `blank` | Пустой слайд с заголовком в стиле шаблона — для свободных элементов `add`. |",
              "| `blank-2` | То же с подзаголовком (заголовок в 2 строки). |", ""]
    with open(os.path.join(SKILL, "reference", "catalog.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    main()
