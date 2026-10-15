#!/usr/bin/env python3
"""Сборка презентаций в фирменном стиле ВШМ СПбГУ из официального шаблона.

Команды:
  list                         — каталог слайдов (ключ, раздел, назначение)
  fields <key> [<key> ...]     — редактируемые фигуры слайда: id, тип, позиция, текст-образец
  build <spec.json> <out.pptx> — собрать презентацию по спецификации (формат: reference/spec.md)
  render <file.pptx> <outdir>  — PNG каждого слайда + contact sheet (нужен LibreOffice)
  check <file.pptx>            — найти неубранные тексты-заглушки шаблона (Lorem, XX, «Название»…)

Как библиотека:
  from gsom_deck import Deck
  deck = Deck(lang="ru"); deck.add("cards-3", {"17": "Заголовок", ...}); deck.save("out.pptx")
"""
import copy
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
ASSETS = os.path.join(SKILL, "assets")
TEMPLATE = os.path.join(ASSETS, "template.pptx")
sys.path.insert(0, HERE)
from catalog import CATALOG, DESCRIPTIONS  # noqa: E402

# ---------------------------------------------------------------- бренд
COLORS = {
    "red": "C31632",        # основной фирменный
    "grey": "4A5761",       # тёмно-серый для презентаций
    "grey2": "6B757D", "grey3": "89929B", "grey4": "ABB0B8", "grey5": "CCCFD4",
    "light": "EBEDF0",      # фон карточек
    "white": "FFFFFF", "black": "000000",
    "gold": "D4A865",       # доп. цвет онлайн-ДПО и MBA/EMBA (заголовки, акценты)
    "bachelor": "006E36", "master": "E68723", "phd": "6E2878",
    "dpo": "013C48", "mba": "0A2463",
}
# Градации для диаграмм (брендбук, с. 27–32): берите соседние оттенки.
PALETTES = {
    "general":  ["C31632", "4A5761", "D94A5E", "89929B", "E8697D", "CCCFD4"],
    "bachelor": ["006E36", "4A5761", "429963", "89929B", "6EBA82", "CCCFD4"],
    "master":   ["E68723", "4A5761", "F0A354", "89929B", "F5B275", "CCCFD4"],
    "phd":      ["6E2878", "4A5761", "8A458F", "89929B", "9C5C9E", "CCCFD4"],
    "dpo":      ["013C48", "4A5761", "1C525C", "89929B", "2F5760", "CCCFD4"],
    "mba":      ["0A2463", "4A5761", "2B4587", "89929B", "405C9E", "CCCFD4"],
    "reds":     ["C31632", "D43B4F", "D94A5E", "E0596E", "E8697D", "F0788C"],
    "greys":    ["4A5761", "6B757D", "89929B", "ABB0B8", "CCCFD4", "EBEDF0"],
}
FONT = "Inter"

# Тексты-заглушки шаблона: если остались в готовом файле — это недоделка.
PLACEHOLDER_PATTERNS = [
    r"lorem", r"ipsum", r"dolor sit", r"consectetur", r"excepteur", r"curabitur", r"quisque",
    r"nulla quis", r"ut enim", r"aenean", r"\bXXX?\b", r"\boption\b", r"название (блока|раздела)",
    r"заголовок\b", r"фамилия", r"имя отчество", r"должность", r"\+7 \(921\)", r"name@gsom",
    r"апрель 2024", r"январь 2024", r"текст цитаты", r"имя автора", r"категория \d", r"ряд \d",
    r"шаг \d", r"\bнаименование\b", r"данные, млрд", r"подзаголовок", r"lorem ipsum",
]

# EN-версия: русский логотип (файл брендбука) -> английский
LOGO_EN = {
    "logo-rus.svg": "logo-eng.svg",
    "logo-partially-inverted-rus.svg": "logo-partially-inverted-eng.svg",
    "logo-grey-rus.svg": "logo-grey-eng.svg",
    "logo-inverted-ru.svg": "logo-inverted-en.svg",
}

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R_ATTRS = ("embed", "link", "id", "pict")


def xp(el, path):
    return etree.ElementBase.xpath(el, path, namespaces=NS)


def hexcolor(c):
    c = COLORS.get(c, c).lstrip("#")
    return RGBColor.from_string(c.upper())


# ---------------------------------------------------------------- обход фигур
def iter_shapes(shapes, depth=0, parent=None):
    for sh in shapes:
        yield sh, depth, parent
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes, depth + 1, sh)


def abs_box(sh):
    """Координаты фигуры на слайде (EMU) с учётом вложенности в группы."""
    if sh.width is None:
        return None
    x, y, w, h = sh.left, sh.top, sh.width, sh.height
    el = sh._element.getparent()
    while el is not None and el.tag == qn("p:grpSp"):
        xf = el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        if xf is None:
            break
        off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
        ch_off, ch_ext = xf.find(qn("a:chOff")), xf.find(qn("a:chExt"))
        kx = int(ext.get("cx")) / int(ch_ext.get("cx")) if int(ch_ext.get("cx")) else 1
        ky = int(ext.get("cy")) / int(ch_ext.get("cy")) if int(ch_ext.get("cy")) else 1
        x = int(off.get("x")) + (x - int(ch_off.get("x"))) * kx
        y = int(off.get("y")) + (y - int(ch_off.get("y"))) * ky
        w, h = w * kx, h * ky
        el = el.getparent()
    return x, y, w, h


def find_shape(slide, sid):
    for sh, _, _ in iter_shapes(slide.shapes):
        if sh.shape_id == sid:
            return sh
    return None


def blipfill_of(sh):
    el = sh._element
    bf = el.find(qn("p:blipFill"))
    if bf is None:
        sppr = el.find(qn("p:spPr"))
        if sppr is not None:
            bf = sppr.find(qn("a:blipFill"))
    return bf


def shape_kind(sh):
    if getattr(sh, "has_chart", False) and sh.has_chart:
        return "chart"
    if getattr(sh, "has_table", False) and sh.has_table:
        return "table"
    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if blipfill_of(sh) is not None:
        return "picture" if sh._element.tag == qn("p:pic") else "image-fill"
    if sh.is_placeholder and sh.placeholder_format.type is not None and \
            "PICTURE" in str(sh.placeholder_format.type):
        return "picture"
    if sh.has_text_frame:
        return "text"
    return "shape"


def slide_key(slide):
    if not slide.has_notes_slide:
        return None
    m = re.match(r"gsom:(\S+)", slide.notes_slide.notes_text_frame.text.strip())
    return m.group(1) if m else None


# ---------------------------------------------------------------- текст
RUN_RE = re.compile(r"(\*\*.+?\*\*)")
# Брендбук, с. 38: не оставлять предлоги и союзы в конце строки -> неразрывный пробел после них
HANGING = re.compile(
    r"(?<![\w-])(и|а|в|к|с|о|у|я|во|ко|со|об|на|не|ни|по|за|от|до|из|без|для|при|над|под|про|"
    r"или|но|что|как|это|И|А|В|К|С|О|У|Я|Во|На|Не|По|За|От|До|Из|Без|Для|При|Над|Под|Про|Но)"
    r" (?=\S)")


def typograph(text):
    text = HANGING.sub(lambda m: m.group(1) + "\u00a0", text)
    return re.sub(r"(\d) (?=(%|₽|руб|тыс|млн|млрд|чел|г\.))", "\\1\u00a0", text)


def _set_runs(p, text, rpr_template):
    """Заполняет абзац текстом; **жирный** -> отдельный run с b=1."""
    for r in list(p):
        if r.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            p.remove(r)
    end = p.find(qn("a:endParaRPr"))
    for chunk in [c for c in RUN_RE.split(text) if c != ""]:
        bold = chunk.startswith("**") and chunk.endswith("**") and len(chunk) > 4
        if bold:
            chunk = chunk[2:-2]
        r = etree.SubElement(p, qn("a:r"))
        if rpr_template is not None:
            rpr = copy.deepcopy(rpr_template)
            rpr.tag = qn("a:rPr")
            r.append(rpr)
        else:
            rpr = etree.SubElement(r, qn("a:rPr"))
        if bold:
            rpr.set("b", "1")
        t = etree.SubElement(r, qn("a:t"))
        t.text = typograph(chunk)
        if end is not None:
            end.addprevious(r)


def _set_lines(p, lines):
    """Один абзац с переносами <a:br/>: i-я строка берёт стиль i-й строки образца."""
    line_styles, cur = [], None
    for child in p:
        if child.tag == qn("a:r") and cur is None:
            cur = child.find(qn("a:rPr"))
        elif child.tag == qn("a:br"):
            line_styles.append(cur)
            cur = None
    line_styles.append(cur)
    br_rpr = next((b.find(qn("a:rPr")) for b in p.iter(qn("a:br"))), None)
    for r in list(p):
        if r.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            p.remove(r)
    end = p.find(qn("a:endParaRPr"))
    for i, line in enumerate(lines):
        if i:
            br = etree.Element(qn("a:br"))
            if br_rpr is not None:
                br.append(copy.deepcopy(br_rpr))
            (end.addprevious(br) if end is not None else p.append(br))
        tmp = etree.Element(qn("a:p"))
        _set_runs(tmp, str(line), line_styles[min(i, len(line_styles) - 1)])
        for r in list(tmp):
            (end.addprevious(r) if end is not None else p.append(r))


def set_text(text_frame, value):
    """value: str ('\\n' — новая строка/абзац) или list[str].
    Форматирование образца сохраняется: i-й абзац (или i-я строка, если в образце один абзац
    с переносами) берёт оформление i-го абзаца/строки образца, лишние — последнего."""
    lines = value if isinstance(value, list) else str(value).split("\n")
    txBody = text_frame._txBody
    paras = txBody.findall(qn("a:p"))
    if len(paras) == 1 and paras[0].find(qn("a:br")) is not None:
        _set_lines(paras[0], lines)
        return
    styles = []
    for p in paras:
        r = p.find(qn("a:r"))
        rpr = r.find(qn("a:rPr")) if r is not None else None
        if rpr is None:
            rpr = p.find(qn("a:endParaRPr"))
        styles.append((p, rpr))
    new_ps = []
    for i, line in enumerate(lines):
        src_p, rpr = styles[min(i, len(styles) - 1)]
        p = copy.deepcopy(src_p)
        _set_runs(p, str(line), rpr)
        new_ps.append(p)
    for p in paras:
        txBody.remove(p)
    for p in new_ps:
        txBody.append(p)


# ---------------------------------------------------------------- клонирование слайдов
def _copy_part(part, package):
    """Глубокая копия части (диаграмма + её xlsx/стили) — у каждого слайда своя."""
    tmpl = re.sub(r"(\d+)(\.\w+)$", r"%d\2", str(part.partname))
    partname = package.next_partname(tmpl)
    cls = type(part)
    if hasattr(cls, "load") and hasattr(part, "_element"):
        new = cls.load(partname, part.content_type, package, part.blob)
    else:
        from pptx.opc.package import Part
        new = Part(partname, part.content_type, package, part.blob)
    mapping = {}
    for rid, rel in part.rels.items():
        if rel.is_external:
            mapping[rid] = new.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            mapping[rid] = new.rels.get_or_add(rel.reltype, _copy_part(rel.target_part, package))
    if hasattr(new, "_element"):
        _remap_rids(new._element, mapping)
    return new


def _remap_rids(root, mapping):
    for el in root.iter():
        for att in R_ATTRS:
            key = f"{{{NS['r']}}}{att}"
            v = el.get(key)
            if v is not None and v in mapping:
                el.set(key, mapping[v])


def clone_slide(prs, src):
    new = prs.slides.add_slide(src.slide_layout)
    sp_tree = new.shapes._spTree
    for el in list(sp_tree):
        if el.tag not in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            sp_tree.remove(el)
    mapping = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            mapping[rid] = new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype in (RT.CHART, RT.PACKAGE, RT.OLE_OBJECT):
            mapping[rid] = new.part.rels.get_or_add(rel.reltype,
                                                    _copy_part(rel.target_part, prs.part.package))
        else:
            mapping[rid] = new.part.rels.get_or_add(rel.reltype, rel.target_part)
    src_tree = src.shapes._spTree
    for el in src_tree:
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        sp_tree.append(copy.deepcopy(el))
    _remap_rids(sp_tree, mapping)
    # фон слайда, если задан на самом слайде
    src_bg = src._element.cSld.find(qn("p:bg"))
    if src_bg is not None:
        bg = copy.deepcopy(src_bg)
        _remap_rids(bg, mapping)
        new._element.cSld.insert(0, bg)
    # переходы/анимации не переносим намеренно
    return new


# ---------------------------------------------------------------- картинки
def _crop_to_fill(blipfill, img_w, img_h, frame_w, frame_h):
    for old in blipfill.findall(qn("a:srcRect")):
        blipfill.remove(old)
    if not (img_w and img_h and frame_w and frame_h):
        return
    img_r, fr_r = img_w / img_h, frame_w / frame_h
    src = etree.Element(qn("a:srcRect"))
    if img_r > fr_r:          # картинка шире — режем бока
        cut = (1 - fr_r / img_r) / 2
        src.set("l", str(int(cut * 100000)))
        src.set("r", str(int(cut * 100000)))
    elif img_r < fr_r:        # выше — режем верх/низ (сверху меньше: лица обычно сверху)
        cut = 1 - img_r / fr_r
        src.set("t", str(int(cut * 0.3 * 100000)))
        src.set("b", str(int(cut * 0.7 * 100000)))
    blip = blipfill.find(qn("a:blip"))
    blip.addnext(src)


def _recolor(el, color):
    rgb = COLORS.get(color, color).lstrip("#").upper()
    for fill in list(el.iter(qn("a:solidFill"))):
        for child in list(fill):
            fill.remove(child)
        etree.SubElement(fill, qn("a:srgbClr"), val=rgb)


def _xfrm(el):
    """xfrm фигуры: у группы — p:grpSpPr/a:xfrm, у остальных — p:spPr/a:xfrm (или p:xfrm)."""
    for path in ("p:grpSpPr/a:xfrm", "p:spPr/a:xfrm", "p:xfrm"):
        found = xp(el, path)
        if found:
            return found[0]
    return None


def put_from(deck, slide, sh, value):
    """{"from": "<key>#<id>", "color": "white"} — взять иконку/иллюстрацию/фото из библиотеки."""
    key, sid = value["from"].split("#")
    lib = deck.source(key)
    src_sh = find_shape(lib, int(sid))
    if src_sh is None:
        raise ValueError(f"{value['from']}: фигура не найдена (см. `fields {key}`)")
    if shape_kind(src_sh) not in ("group", "picture", "image-fill", "shape", "text") or src_sh.is_placeholder:
        raise ValueError(f"{value['from']}: источником может быть иконка (group) или картинка")
    src_bf, dst_bf = blipfill_of(src_sh), blipfill_of(sh)
    is_photo = src_sh._element.tag == qn("p:pic") and not key.startswith("lib-icons")
    if is_photo and dst_bf is not None:           # фото -> рамка фото: меняем только заливку
        mapping = {}
        for rid in xp(src_bf, ".//@r:embed | .//@r:link"):
            rel = lib.part.rels[rid]
            mapping[rid] = slide.part.rels.get_or_add(rel.reltype, rel.target_part)
        new_bf = copy.deepcopy(src_bf)
        _remap_rids(new_bf, mapping)
        new_bf.tag = dst_bf.tag
        dst_bf.getparent().replace(dst_bf, new_bf)
        img = lib.part.rels[xp(src_bf, ".//a:blip/@r:embed")[0]].target_part
        if hasattr(img, "image"):
            _crop_to_fill(new_bf, *img.image.size, sh.width, sh.height)
        return
    # иконка/иллюстрация: копия фигуры, вписанная по центру в рамку заменяемой
    new_el = copy.deepcopy(src_sh._element)
    mapping = {}
    for rid in xp(new_el, ".//@r:embed | .//@r:link | .//@r:id"):
        if rid in lib.part.rels:
            rel = lib.part.rels[rid]
            mapping[rid] = slide.part.rels.get_or_add(rel.reltype, rel.target_part)
    _remap_rids(new_el, mapping)
    if value.get("color"):
        _recolor(new_el, value["color"])
    sx = _xfrm(new_el)
    s_ext = sx.find(qn("a:ext"))
    bx, by, bw, bh = int(sh.left), int(sh.top), int(sh.width), int(sh.height)
    sw, sh_ = int(s_ext.get("cx")), int(s_ext.get("cy"))
    k = min(bw / sw, bh / sh_) if sw and sh_ else 1
    nw, nh = int(sw * k), int(sh_ * k)
    sx.find(qn("a:off")).set("x", str(bx + (bw - nw) // 2))
    sx.find(qn("a:off")).set("y", str(by + (bh - nh) // 2))
    s_ext.set("cx", str(nw)); s_ext.set("cy", str(nh))
    used = {int(v) for v in xp(slide.shapes._spTree, ".//p:cNvPr/@id")}
    next_id = max(used | {0}) + 1
    for i, nv in enumerate(new_el.iter(qn("p:cNvPr"))):
        if i == 0:
            nv.set("id", str(sh.shape_id))  # id сохраняем — фигуру можно адресовать дальше
        else:
            nv.set("id", str(next_id)); next_id += 1
    sh._element.getparent().replace(sh._element, new_el)


def _grayscale(path):
    """Ч/б версия фото (брендбук, с. 68: для онлайн-ДПО и MBA/EMBA — только ч/б)."""
    from PIL import Image, ImageOps
    out = os.path.join(tempfile.mkdtemp(prefix="gsom_"), "bw_" + os.path.basename(path))
    im = Image.open(path)
    im = ImageOps.exif_transpose(im).convert("L")
    if out.lower().endswith((".jpg", ".jpeg")):
        im.save(out, quality=90)
    else:
        im.save(out)
    return out


def set_picture(deck, slide, sh, value):
    if isinstance(value, dict) and "from" in value:
        return put_from(deck, slide, sh, value)
    bf = blipfill_of(sh)
    path = value["path"] if isinstance(value, dict) else value
    path = os.path.expanduser(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    if isinstance(value, dict) and value.get("grayscale"):
        path = _grayscale(path)
    if bf is None:   # пустой плейсхолдер картинки
        sh.insert_picture(path)
        return
    image_part, rid = slide.part.get_or_add_image_part(path)
    blip = bf.find(qn("a:blip"))
    blip.set(qn("r:embed"), rid)
    for ext_lst in blip.findall(qn("a:extLst")):     # убрать svg-версию образца
        for ext in list(ext_lst):
            if "96DAC541" in (ext.get("uri") or ""):
                ext_lst.remove(ext)
    if not (isinstance(value, dict) and value.get("fit") == "stretch"):
        w, h = image_part.image.size
        _crop_to_fill(bf, w, h, sh.width, sh.height)


# ---------------------------------------------------------------- таблицы и диаграммы
def set_table(slide, tbl_shape, value):
    """value: list строк или {"rows": [...], "total": true} — последняя строка берёт
    оформление последней строки образца («Итого»)."""
    rows = value["rows"] if isinstance(value, dict) else value
    keep_total = isinstance(value, dict) and value.get("total")
    tbl = tbl_shape.table
    old_box = (tbl_shape.left, tbl_shape.top, tbl_shape.width, tbl_shape.height)
    tr_list = tbl._tbl.tr_lst
    if keep_total and len(tr_list) > 2:
        total_tr = tr_list[-1]
        tbl._tbl.remove(total_tr)
        while len(tbl._tbl.tr_lst) > max(len(rows) - 1, 1):
            tbl._tbl.remove(tbl._tbl.tr_lst[-1])
        while len(tbl._tbl.tr_lst) < len(rows) - 1:
            tbl._tbl.tr_lst[-1].addnext(copy.deepcopy(tbl._tbl.tr_lst[-1]))
        tbl._tbl.tr_lst[-1].addnext(total_tr)
    tr_list = tbl._tbl.tr_lst
    while len(tr_list) < len(rows):              # добавить строки по образцу последней
        new_tr = copy.deepcopy(tr_list[-1])
        tr_list[-1].addnext(new_tr)
        tr_list = tbl._tbl.tr_lst
    for tr in tr_list[len(rows):]:                # лишние строки удалить
        tbl._tbl.remove(tr)
    total_h = 0
    for r, row in enumerate(rows):
        for c in range(len(tbl.columns)):
            val = row[c] if c < len(row) else ""
            if val is None:
                continue
            cell = tbl.cell(r, c)
            if cell.is_spanned:
                continue
            set_text(cell.text_frame, val)
        total_h += tbl.rows[r].height
    tbl_shape.height = Emu(total_h)
    # подложка таблицы (фигура/группа точно под таблицей) — подогнать по высоте
    tol = Inches(0.06)
    for sh in slide.shapes:
        if sh.shape_id == tbl_shape.shape_id or sh.width is None:
            continue
        if all(abs(a - b) < tol for a, b in zip((sh.left, sh.top, sh.width, sh.height), old_box)):
            _resize_height(sh, old_box[3], total_h)


def _resize_height(sh, old_h, new_h):
    """Подложка: у группы меняем только «полновысотные» элементы, шапки и линии не трогаем."""
    if sh.shape_type != MSO_SHAPE_TYPE.GROUP:
        sh.height = Emu(new_h)
        return
    xfrm = _xfrm(sh._element)
    ext, ch_ext, ch_off = xfrm.find(qn("a:ext")), xfrm.find(qn("a:chExt")), xfrm.find(qn("a:chOff"))
    k = int(ch_ext.get("cy")) / int(ext.get("cy")) if int(ext.get("cy")) else 1
    old_c, new_c = old_h * k, new_h * k
    top = int(ch_off.get("y"))
    for child in sh.shapes:
        cx = _xfrm(child._element)
        if cx is None:
            continue
        c_off, c_ext = cx.find(qn("a:off")), cx.find(qn("a:ext"))
        cy, ch = int(c_off.get("y")), int(c_ext.get("cy"))
        if abs(ch - old_c) < old_c * 0.03:                 # во всю высоту — растянуть/сжать
            c_ext.set("cy", str(int(new_c)))
        elif cy + ch > top + old_c * 0.9:                  # прижата к низу — сдвинуть
            c_off.set("y", str(int(cy + new_c - old_c)))
    ext.set("cy", str(int(new_h)))
    ch_ext.set("cy", str(int(new_c)))


def set_chart(chart_shape, value):
    chart = chart_shape.chart
    data = CategoryChartData()
    data.categories = value["categories"]
    for s in value["series"]:
        data.add_series(s["name"], s["values"], number_format=s.get("format"))
    chart.replace_data(data)
    if "legend" in value:
        chart.has_legend = bool(value["legend"])
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
    # масштаб осей образца (например, min=500) под новые данные не подходит — сбрасываем
    for scaling in chart._chartSpace.iter(qn("c:scaling")):
        for tag in ("c:min", "c:max"):
            for el in scaling.findall(qn(tag)):
                scaling.remove(el)
    for tag in ("c:majorUnit", "c:minorUnit"):
        for el in list(chart._chartSpace.iter(qn(tag))):
            el.getparent().remove(el)
    for ax in (chart.value_axis,) if value.get("min") is not None or value.get("max") is not None else ():
        if value.get("min") is not None:
            ax.minimum_scale = value["min"]
        if value.get("max") is not None:
            ax.maximum_scale = value["max"]


# ---------------------------------------------------------------- свободные элементы
def _box(slide, spec):
    return [Inches(spec[k]) for k in ("x", "y", "w", "h")]


def _style_runs(tf, size, bold, color, align):
    for p in tf.paragraphs:
        if align:
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER}[align]
        for r in p.runs:
            r.font.name = FONT
            r.font.size = Pt(size)
            if bold:
                r.font.bold = True
            r.font.color.rgb = hexcolor(color)


def add_text(slide, spec):
    x, y, w, h = _box(slide, spec)
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = spec["text"] if isinstance(spec["text"], list) else str(spec["text"]).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _set_runs(p._p, ("• " + line) if spec.get("bullets") else line, None)
        p.space_after = Pt(spec.get("space_after", 6))
    _style_runs(tf, spec.get("size", 14), spec.get("bold"), spec.get("color", "grey"),
                spec.get("align"))
    for r in tf._txBody.iter(qn("a:r")):     # **жирный** внутри текста
        rpr = r.find(qn("a:rPr"))
        if rpr is not None and rpr.get("b") == "1":
            continue
    return tb


def add_card(slide, spec):
    """Карточка как в шаблоне: светло-серое тело + шапка (красная/серая) с белым текстом."""
    x, y, w, h = _box(slide, spec)
    head_h = Inches(spec.get("head_h", 0.6))
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    body.adjustments[0] = 0.03
    body.fill.solid(); body.fill.fore_color.rgb = hexcolor("light")
    body.line.fill.background(); body.shadow.inherit = False
    head = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, head_h)
    head.adjustments[0] = 0.1
    head.fill.solid(); head.fill.fore_color.rgb = hexcolor(spec.get("accent", "red"))
    head.line.fill.background(); head.shadow.inherit = False
    tf = head.text_frame
    tf.margin_left = tf.margin_right = Inches(0.15)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    _set_runs(tf.paragraphs[0]._p, spec.get("title", ""), None)
    _style_runs(tf, spec.get("title_size", 13), True, "white", None)
    if spec.get("text"):
        add_text(slide, {"x": spec["x"] + 0.2, "y": spec["y"] + spec.get("head_h", 0.6) + 0.2,
                         "w": spec["w"] - 0.4, "h": spec["h"] - spec.get("head_h", 0.6) - 0.3,
                         "text": spec["text"], "size": spec.get("size", 12),
                         "bullets": spec.get("bullets"), "color": "grey"})
    return body


def add_table(slide, spec):
    rows = spec["rows"]
    x, y, w, h = _box(slide, spec)
    shape = slide.shapes.add_table(len(rows), max(len(r) for r in rows), x, y, w, h)
    tbl = shape.table
    tbl.first_row = True
    tbl_pr = tbl._tbl.tblPr
    style = tbl_pr.find(qn("a:tableStyleId"))
    if style is not None:
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # «No Style, Table Grid»
    size = spec.get("size", 11)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = "" if val is None else str(val)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            head = r == 0
            total = spec.get("total_row") and r == len(rows) - 1
            cell.fill.solid()
            cell.fill.fore_color.rgb = hexcolor("red" if head else ("light" if total or r % 2 == 0 else "white"))
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.name = FONT
                    run.font.size = Pt(size)
                    run.font.bold = head or bool(total)
                    run.font.color.rgb = hexcolor("white" if head else ("red" if total else "grey"))
    return shape


CHART_TYPES = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column-stacked": XL_CHART_TYPE.COLUMN_STACKED, "bar-stacked": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}


def add_chart(slide, spec):
    x, y, w, h = _box(slide, spec)
    data = CategoryChartData()
    data.categories = spec["categories"]
    for s in spec["series"]:
        data.add_series(s["name"], s["values"])
    kind = spec.get("chart", "column")
    gf = slide.shapes.add_chart(CHART_TYPES[kind], x, y, w, h, data)
    chart = gf.chart
    chart.has_title = bool(spec.get("title"))
    if spec.get("title"):
        chart.chart_title.text_frame.text = spec["title"]
    pal = PALETTES[spec.get("palette", "general")]
    chart.font.name = FONT
    chart.font.size = Pt(spec.get("size", 11))
    chart.font.color.rgb = hexcolor("grey")
    if kind in ("pie", "doughnut"):
        plot = chart.plots[0]
        for i, pt in enumerate(plot.series[0].points):
            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = hexcolor(pal[i % len(pal)])
    else:
        for i, s in enumerate(chart.plots[0].series):
            col = hexcolor(pal[i % len(pal)])
            if kind == "line":
                s.format.line.color.rgb = col
                s.format.line.width = Pt(2.25)
                s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = col
                s.marker.format.line.color.rgb = col
            else:
                s.format.fill.solid(); s.format.fill.fore_color.rgb = col
        if kind not in ("line",):
            chart.plots[0].gap_width = 60
        va = chart.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = hexcolor("light")
        va.format.line.fill.background()
        chart.category_axis.format.line.color.rgb = hexcolor("grey5")
    if spec.get("labels", True):
        plot = chart.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(spec.get("size", 11))
        dl.font.name = FONT
        if spec.get("format"):
            dl.number_format = spec["format"]; dl.number_format_is_linked = False
    multi = len(spec["series"]) > 1 or kind in ("pie", "doughnut")
    chart.has_legend = spec.get("legend", multi)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(spec.get("size", 11))
    return gf


CONTAINER = [(0.0, 0.447), (0.165, 0.0), (1.0, 0.0), (1.0, 0.553), (0.835, 1.0), (0.0, 1.0)]


def add_image(slide, spec):
    """Фото; container=true — фирменный контейнер (прямоугольник с двумя срезанными углами,
    брендбук с. 65–67). Наклон среза постоянный: −26.61°."""
    path = os.path.expanduser(spec["path"])
    if spec.get("grayscale"):
        path = _grayscale(path)
    x, y, w, h = _box(slide, spec)
    pic = slide.shapes.add_picture(path, x, y, w, h)
    bf = pic._element.find(qn("p:blipFill"))
    _crop_to_fill(bf, *pic.image.size, w, h)
    if spec.get("container"):
        # срез: по горизонтали dx = h_cut / tan(26.61°); берём h_cut = 0.447*h
        import math
        hc = 0.447 * h
        dx = min(hc / math.tan(math.radians(63.39)), w * 0.3)
        pts = [(0, hc), (dx, 0), (w, 0), (w, h - hc), (w - dx, h), (0, h)]
        sppr = pic._element.spPr
        prst = sppr.find(qn("a:prstGeom"))
        cust = etree.SubElement(sppr, qn("a:custGeom"))
        for tag in ("avLst", "gdLst", "ahLst", "cxnLst"):
            etree.SubElement(cust, qn(f"a:{tag}"))
        etree.SubElement(cust, qn("a:rect"), l="0", t="0", r="r", b="b")
        lst = etree.SubElement(cust, qn("a:pathLst"))
        path_el = etree.SubElement(lst, qn("a:path"), w=str(int(w)), h=str(int(h)))
        for i, (px, py) in enumerate(pts):
            cmd = etree.SubElement(path_el, qn("a:moveTo" if i == 0 else "a:lnTo"))
            etree.SubElement(cmd, qn("a:pt"), x=str(int(px)), y=str(int(py)))
        etree.SubElement(path_el, qn("a:close"))
        if prst is not None:
            prst.addprevious(cust)
            sppr.remove(prst)
    return pic


ADDERS = {"text": add_text, "card": add_card, "table": add_table, "chart": add_chart,
          "image": add_image}


# ---------------------------------------------------------------- колода
class Deck:
    def __init__(self, lang="ru", template=TEMPLATE):
        self.prs = Presentation(template)
        self.lang = lang
        self._sources = {}
        self._originals = list(self.prs.slides)
        for s in self._originals:
            k = slide_key(s)
            if k:
                self._sources[k] = s
        self._layouts = {l.name: l for l in self.prs.slide_masters[0].slide_layouts}
        if lang == "en":
            self._english_logos()

    # --- источники
    def keys(self):
        return list(self._sources)

    def source(self, key):
        if key not in self._sources:
            raise KeyError(f"нет слайда «{key}» в каталоге (см. `gsom_deck.py list`)")
        return self._sources[key]

    def _english_logos(self):
        logo_dir = os.path.join(ASSETS, "logos")
        ru = {}
        for ru_name, en_name in LOGO_EN.items():
            b = open(os.path.join(logo_dir, ru_name), "rb").read()
            ru[hashlib.md5(b).hexdigest()] = open(os.path.join(logo_dir, en_name), "rb").read()
        aspects = {}
        for part in self.prs.part.package.iter_parts():
            if part.content_type == "image/svg+xml":
                new = ru.get(hashlib.md5(part.blob).hexdigest())
                if new:
                    part._blob = new
                    m = re.search(rb'viewBox="\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)', new)
                    aspects[part.partname] = float(m.group(1)) / float(m.group(2))
        holders = list(self.prs.slide_masters)
        holders += [l for m in self.prs.slide_masters for l in m.slide_layouts] + self._originals
        for holder in holders:
            rels = holder.part.rels
            for pic in holder.shapes._spTree.iter(qn("p:pic")):
                aspect = None
                for rid in xp(pic, ".//@r:embed"):
                    if rid in rels and not rels[rid].is_external:
                        aspect = aspects.get(rels[rid].target_part.partname, aspect)
                if aspect is None:
                    continue
                off, ext = pic.spPr.find(qn("a:xfrm")).find(qn("a:off")), \
                    pic.spPr.find(qn("a:xfrm")).find(qn("a:ext"))
                x, cx, cy = int(off.get("x")), int(ext.get("cx")), int(ext.get("cy"))
                new_cx = int(cy * aspect)
                if x + cx / 2 > self.prs.slide_width / 2:
                    off.set("x", str(x + cx - new_cx))
                ext.set("cx", str(new_cx))

    # --- добавление слайдов
    def add(self, key, fields=None, add=None, notes=None):
        if key in ("blank", "blank-2"):
            # клон контентного слайда без содержимого: заголовок сохраняет оформление шаблона
            slide = clone_slide(self.prs, self.source("bullets-4" if key == "blank" else "lib-icons-people"))
            for sh in list(slide.shapes):
                if not sh.is_placeholder:
                    sh._element.getparent().remove(sh._element)
        else:
            slide = clone_slide(self.prs, self.source(key))
        self.fill(slide, fields or {})
        for spec in add or []:
            ADDERS[spec["type"]](slide, spec)
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        return slide

    def fill(self, slide, fields):
        deletions = []
        for fid, value in fields.items():
            if fid == "title":
                sh = next((s for s in slide.placeholders if s.placeholder_format.idx == 0), None)
                if sh is None:
                    raise ValueError("на слайде нет плейсхолдера заголовка")
            elif fid.startswith("ph"):
                idx = int(fid[2:])
                sh = next((s for s in slide.placeholders if s.placeholder_format.idx == idx), None)
            else:
                sh = find_shape(slide, int(fid.lstrip("#")))
            if sh is None:
                raise ValueError(f"фигура {fid} не найдена (см. `gsom_deck.py fields <key>`)")
            if value is None:
                deletions.append(sh)
                continue
            kind = shape_kind(sh)
            if isinstance(value, dict) and "from" in value:
                put_from(self, slide, sh, value)
            elif kind == "table":
                set_table(slide, sh, value)
            elif kind == "chart":
                set_chart(sh, value)
            elif kind in ("picture", "image-fill"):
                set_picture(self, slide, sh, value)
            elif sh.has_text_frame:
                set_text(sh.text_frame, value)
            else:
                raise ValueError(f"фигура {fid} ({kind}) не принимает значение")
        for sh in deletions:
            el = sh._element
            el.getparent().remove(el)

    def save(self, path):
        sld_ids = self.prs.slides._sldIdLst
        originals = {id(s.part) for s in self._originals}
        for sld_id in list(sld_ids):
            part = self.prs.part.related_part(sld_id.rId)
            if id(part) in originals:
                self.prs.part.drop_rel(sld_id.rId)
                sld_ids.remove(sld_id)
        for n, slide in enumerate(self.prs.slides, 1):   # кэш номеров слайдов
            for fld in slide.shapes._spTree.iter(qn("a:fld")):
                if fld.get("type") == "slidenum":
                    t = fld.find(qn("a:t"))
                    if t is not None:
                        t.text = str(n)
        self.prs.save(path)
        return path

    def overflow(self):
        """Фигуры, выходящие за край слайда (частая причина: слишком много строк в таблице)."""
        out = []
        W, H = self.prs.slide_width, self.prs.slide_height
        tol = Inches(0.05)
        originals = {id(s.part) for s in self._originals}
        n = 0
        for slide in self.prs.slides:
            if id(slide.part) in originals:
                continue
            n += 1
            for sh in slide.shapes:
                if sh.width is None or sh.is_placeholder:
                    continue
                if sh.left + sh.width > W + tol or sh.top + sh.height > H + tol:
                    out.append(f"слайд {n}, фигура #{sh.shape_id}: выходит за край слайда")
        return out


# ---------------------------------------------------------------- команды
def cmd_list():
    section = None
    for key in CATALOG.values():
        sec, desc = DESCRIPTIONS.get(key, ("?", ""))
        if sec != section:
            print(f"\n## {sec}")
            section = sec
        print(f"  {key:26} {desc}")
    print("\n  blank                      Пустой слайд с заголовком (для свободных элементов add).")
    print("  blank-2                    То же, заголовок в 2+ строки.")


def describe(sh, slide):
    kind = shape_kind(sh)
    box = abs_box(sh)
    pos = "x={:.2f} y={:.2f} w={:.2f} h={:.2f}".format(*(Emu(int(v)).inches for v in box)) \
        if box else ""
    extra = ""
    if kind == "table":
        t = sh.table
        sample = [[t.cell(r, c).text.replace("\n", " / ")[:18] for c in range(len(t.columns))]
                  for r in range(min(3, len(t.rows)))]
        extra = f"{len(t.rows)}×{len(t.columns)} " + json.dumps(sample, ensure_ascii=False)
    elif kind == "chart":
        ch = sh.chart
        extra = f"{ch.chart_type} кат.={list(ch.plots[0].categories)[:6]} рядов={len(ch.plots[0].series)}"
    elif sh.has_text_frame:
        txt = sh.text_frame.text.replace("\n", " ⏎ ")
        n = len(sh.text_frame.paragraphs)
        extra = (f"[{n} абз.] " if n > 1 else "") + (f"«{txt[:70]}»" if txt.strip() else "(пусто)")
    ph = f" ph{sh.placeholder_format.idx}" if sh.is_placeholder else ""
    return f"{kind}{ph} {pos} {extra}"


def fillable(slide):
    """(shape, depth) фигур, которые имеет смысл заполнять/удалять — как в `fields`.
    В библиотеках (lib-*) — все ассеты верхнего уровня (иконки, иллюстрации, фото)."""
    if (slide_key(slide) or "").startswith("lib-"):
        for sh in slide.shapes:
            if not sh.is_placeholder:
                yield sh, 0
        return
    for sh, depth, _ in iter_shapes(slide.shapes):
        if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
            continue
        kind = shape_kind(sh)
        empty = not (sh.has_text_frame and sh.text_frame.text.strip())
        if empty and (kind == "shape" or (kind == "text" and not sh.is_placeholder)):
            continue
        yield sh, depth


def cmd_fields(keys):
    deck = Deck()
    for key in keys:
        slide = deck.source(key)
        print(f"\n=== {key} — {DESCRIPTIONS.get(key, ('', ''))[1]}")
        print(f"    layout: {slide.slide_layout.name}")
        print(f"    превью с номерами фигур: assets/catalog/{key}.jpg")
        for sh, depth in fillable(slide):
            print("    " + "  " * depth + f"#{sh.shape_id:<4} " + describe(sh, slide))


def cmd_build(spec_path, out):
    spec = json.load(open(spec_path, encoding="utf-8"))
    base = os.path.dirname(os.path.abspath(spec_path))
    deck = Deck(lang=spec.get("lang", "ru"))
    for i, s in enumerate(spec["slides"], 1):
        fields = s.get("fields", {})
        for k, v in list(fields.items()):          # относительные пути картинок — от spec.json
            if isinstance(v, str) and re.search(r"\.(png|jpe?g|gif|bmp|tiff?)$", v, re.I) \
                    and not os.path.isabs(v):
                cand = os.path.join(base, v)
                if os.path.exists(cand):
                    fields[k] = cand
        for a in s.get("add", []):
            if a.get("type") == "image" and not os.path.isabs(a["path"]):
                a["path"] = os.path.join(base, a["path"])
        try:
            deck.add(s["use"], fields, s.get("add"), s.get("notes"))
        except Exception as e:
            sys.exit(f"Слайд {i} ({s.get('use')}): {e}")
    over = deck.overflow()
    deck.save(out)
    print(f"OK: {len(spec['slides'])} слайдов -> {out}")
    if over:
        print("\nВНИМАНИЕ: " + "\n  ".join(["элементы за краем слайда:"] + over))
    problems = find_placeholders(out)
    if problems:
        print("\nВНИМАНИЕ: остались тексты-заглушки шаблона (заполните или удалите через null):")
        for p in problems:
            print("  " + p)


def find_placeholders(path):
    prs = Presentation(path)
    rx = re.compile("|".join(PLACEHOLDER_PATTERNS), re.I)
    out = []
    for n, slide in enumerate(prs.slides, 1):
        for sh, _, _ in iter_shapes(slide.shapes):
            texts = []
            if getattr(sh, "has_table", False) and sh.has_table:
                texts = [c.text for row in sh.table.rows for c in row.cells]
            elif sh.has_text_frame:
                texts = [sh.text_frame.text]
            for t in texts:
                m = rx.search(t)
                if m:
                    out.append(f"слайд {n}, фигура #{sh.shape_id}: «{t.strip()[:60]}»")
                    break
    return out


def ensure_fonts():
    try:
        have = subprocess.run(["fc-list", ":family=Inter"], capture_output=True, text=True).stdout
    except FileNotFoundError:
        return
    if "Inter" in have:
        return
    dst = os.path.expanduser("~/.fonts")
    os.makedirs(dst, exist_ok=True)
    for f in glob.glob(os.path.join(ASSETS, "fonts", "*.ttf")):
        shutil.copy(f, dst)
    subprocess.run(["fc-cache", "-f"], capture_output=True)


def cmd_render(pptx, outdir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("Нужен LibreOffice (soffice) для рендера.")
    ensure_fonts()
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp,
                        os.path.abspath(pptx)], capture_output=True, check=True)
        pdf = glob.glob(os.path.join(tmp, "*.pdf"))[0]
        for old in glob.glob(os.path.join(outdir, "slide-*.png")):
            os.remove(old)
        subprocess.run(["pdftoppm", "-r", "80", "-png", pdf, os.path.join(outdir, "slide")],
                       check=True)
    pngs = sorted(glob.glob(os.path.join(outdir, "slide-*.png")))
    try:
        from PIL import Image, ImageDraw
        ims = [Image.open(p) for p in pngs]
        w, h = ims[0].size
        scale = 0.45
        tw, th = int(w * scale), int(h * scale)
        cols = 3
        rows = (len(ims) + cols - 1) // cols
        sheet = Image.new("RGB", (cols * tw + (cols + 1) * 8, rows * (th + 22) + 8), "white")
        d = ImageDraw.Draw(sheet)
        for i, im in enumerate(ims):
            x = 8 + (i % cols) * (tw + 8)
            y = 8 + (i // cols) * (th + 22)
            d.text((x, y), str(i + 1), fill="black")
            sheet.paste(im.resize((tw, th)), (x, y + 14))
        sheet.save(os.path.join(outdir, "contact-sheet.png"))
    except Exception as e:  # PIL не обязателен
        print("contact sheet не собран:", e)
    print(f"{len(pngs)} слайдов -> {outdir}/slide-*.png, {outdir}/contact-sheet.png")


def main(argv):
    if len(argv) < 2:
        sys.exit(__doc__)
    cmd = argv[1]
    if cmd == "list":
        cmd_list()
    elif cmd == "fields" and len(argv) > 2:
        cmd_fields(argv[2:])
    elif cmd == "build" and len(argv) == 4:
        cmd_build(argv[2], argv[3])
    elif cmd == "render" and len(argv) == 4:
        cmd_render(argv[2], argv[3])
    elif cmd == "check" and len(argv) == 3:
        problems = find_placeholders(argv[2])
        print("\n".join(problems) if problems else "OK: заглушек шаблона не найдено")
        sys.exit(1 if problems else 0)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv)
