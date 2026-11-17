#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Массовый перенос презентаций в шаблон УГМУ.

Дизайн шаблона УГМУ нарисован прямо на слайдах шаблона (бордовые титульные слайды,
шапка с названием раздела, линией и номером, логотип, дуги). Поэтому скрипт берёт
из шаблона три готовые "рамки" и наполняет их содержимым исходных слайдов:

  1-й слайд        -> рамка титульного слайда (шаблон, слайд 5)
  слайд из одного  -> рамка раздела           (шаблон, слайд 6)
    заголовка
  все остальные    -> рамка со шапкой         (шаблон, слайд 2)

Всё содержимое слайда (текст, картинки, таблицы, диаграммы, группы, фигуры, медиа,
гиперссылки, анимации, заметки) копируется целиком и вписывается в рабочую область
под заголовком. После конвертации проверяется, что слова/картинки/таблицы/диаграммы
не потерялись; результат пишется в report.csv.

Установка:  pip install python-pptx lxml
Проверка:   python convert_to_ugmu.py --template template.pptx --inspect
Запуск:     python convert_to_ugmu.py --template template.pptx --src old --out new
"""
import argparse, copy, csv, re, sys, traceback
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER as PH
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.shapes.shapetree import SlideShapeFactory
from pptx.util import Emu

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS, "r": R_NS}

# ---- Сетка шаблона (EMU). Взята из слайдов шаблона УГМУ ----
LEFT, RIGHT = 1043364, 11659949       # поля текстовых блоков
TITLE_TOP = 830000                     # заголовок на слайдах со шапкой
GAP = 200000                           # отступ заголовок -> контент
BOTTOM_MARGIN = 380000
LINE_END = 10787605                    # правый конец линии в шапке
BODY_FONT = "Montserrat Light"
BOLD_FONT = "Montserrat SemiBold"
SYMBOL_FONTS = ("wingdings", "symbol", "webdings", "marlett")
TITLE_TYPES = (PH.TITLE, PH.CENTER_TITLE)
DROP_PH = (PH.DATE, PH.FOOTER, PH.SLIDE_NUMBER)
SKIP_TAGS = {qn("p:nvGrpSpPr"), qn("p:grpSpPr"), qn("p:extLst")}


# ================================================================ утилиты
def sp_text(el):
    return "".join(t.text or "" for t in el.iter(qn("a:t")))


def paras(el):
    """Абзацы текста; разрыв строки (Shift+Enter) заменяется пробелом, чтобы слова не склеивались."""
    out = []
    for p in el.iter(qn("a:p")):
        parts = []
        for ch in p:
            if ch.tag == qn("a:br"):
                parts.append(" ")
            else:
                parts.append("".join(x.text or "" for x in ch.iter(qn("a:t"))))
        t = re.sub(r"\s+", " ", "".join(parts)).strip()
        if t:
            out.append(t)
    return out


def n_lines(text, width_emu, pt, cw):
    per = max(1, int(width_emu / 914400 * 72 / (pt * cw)))
    lines, cur = 1, 0
    for w in text.split():
        add = len(w) + (1 if cur else 0)
        if cur + add > per and cur:
            lines += 1; cur = len(w)
        else:
            cur += add
    return lines


def fit_pt(text, width_emu, max_h, base, lnsp, cw, min_pt):
    pt = base
    while pt > min_pt:
        h = n_lines(text, width_emu, pt, cw) * pt * 1.2 * lnsp / 72 * 914400
        if h <= max_h:
            break
        pt -= 1
    lines = n_lines(text, width_emu, pt, cw)
    return pt, lines, lines * pt * 1.2 * lnsp / 72 * 914400


def run_props(txbody_sp):
    """Первый rPr, размер и межстрочный интервал шаблонного текстового блока."""
    r = txbody_sp.find(".//a:r/a:rPr", NS)
    sz = int(r.get("sz")) / 100 if r is not None and r.get("sz") else 18
    ln = txbody_sp.find(".//a:pPr/a:lnSpc/a:spcPct", NS)
    return sz, (int(ln.get("val")) / 100000 if ln is not None else 0.9)


def fill_shape(sp, lines, pt=None):
    """Подставляет текст в шаблонный блок, сохраняя его оформление."""
    tb = sp.find(qn("p:txBody"))
    ps = tb.findall(qn("a:p"))
    tp = ps[0]
    r0 = tp.find(qn("a:r"))
    rpr0 = r0.find(qn("a:rPr")) if r0 is not None else None
    for p in ps:
        tb.remove(p)
    for line in lines:
        p = copy.deepcopy(tp)
        for ch in list(p):
            if ch.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
                p.remove(ch)
        r = etree.Element(qn("a:r"))
        if rpr0 is not None:
            rp = copy.deepcopy(rpr0)
            if pt:
                rp.set("sz", str(int(pt * 100)))
            r.append(rp)
        etree.SubElement(r, qn("a:t")).text = line
        epr = p.find(qn("a:endParaRPr"))
        if epr is not None:
            epr.addprevious(r)
            if pt:
                epr.set("sz", str(int(pt * 100)))
        else:
            p.append(r)
        tb.append(p)


def set_xfrm(el, x=None, y=None, cx=None, cy=None):
    xf = el.find(".//a:xfrm", NS)
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    if x is not None: off.set("x", str(int(x)))
    if y is not None: off.set("y", str(int(y)))
    if cx is not None: ext.set("cx", str(int(cx)))
    if cy is not None: ext.set("cy", str(int(cy)))


def get_xfrm(el):
    xf = el.find(".//a:xfrm", NS)
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


# ================================================================ копирование частей пакета
def _tmpl(partname):
    s = str(partname)
    m = re.search(r"(\d+)(\.[^./]+)$", s)
    if m:
        return s[:m.start(1)] + "%d" + m.group(2)
    stem, dot, ext = s.rpartition(".")
    return stem + "%d" + dot + ext


def copy_part(src_part, package, memo):
    if id(src_part) in memo:
        return memo[id(src_part)]
    name = package.next_partname(_tmpl(src_part.partname))
    new = type(src_part).load(PackURI(name), src_part.content_type, package, src_part.blob)
    memo[id(src_part)] = new
    return new


def copy_rels(src_part, dst_part, package, memo, deferred, skip_types, same_package=False):
    rid_map = {}
    for rid, rel in src_part.rels.items():
        if rel.reltype in skip_types:
            continue
        if rel.reltype == RT.SLIDE:
            deferred.append((dst_part, rid, rel.target_part))
            continue
        if rel.is_external:
            rid_map[rid] = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif same_package:
            rid_map[rid] = dst_part.relate_to(rel.target_part, rel.reltype)
        else:
            tgt = rel.target_part
            fresh = id(tgt) not in memo
            new = copy_part(tgt, package, memo)
            rid_map[rid] = dst_part.relate_to(new, rel.reltype)
            if fresh:
                copy_rels(tgt, new, package, memo, deferred, skip_types=())
    return rid_map


def remap_rids(root, rid_map):
    pre = "{" + R_NS + "}"
    for el in root.iter():
        for k, v in list(el.attrib.items()):
            if k.startswith(pre) and v in rid_map:
                el.set(k, rid_map[v])


# ================================================================ шрифты и масштаб
def set_font(rpr, face):
    cur = rpr.find(qn("a:latin"))
    if cur is not None and cur.get("typeface", "").lower().startswith(SYMBOL_FONTS):
        return
    for tag in ("latin", "cs"):
        for ch in rpr.findall(qn("a:" + tag)):
            rpr.remove(ch)
    latin = etree.Element(qn("a:latin")); latin.set("typeface", face)
    cs = etree.Element(qn("a:cs")); cs.set("typeface", "Montserrat")
    anchor = None
    for tag in ("ea", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"):
        anchor = rpr.find(qn("a:" + tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(latin)
        anchor.addprevious(cs) if anchor.tag != qn("a:ea") else anchor.addnext(cs)
    else:
        rpr.append(latin); rpr.append(cs)


def restyle_fonts(el):
    """Единый шрифт шаблона: Montserrat Light, жирный - SemiBold. Возвращает True, если шрифты менялись."""
    changed = False
    for r in el.iter(qn("a:r"), qn("a:fld")):
        rpr = r.find(qn("a:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("a:rPr")); rpr.set("lang", "ru-RU")
            r.insert(0, rpr)
        lat = rpr.find(qn("a:latin"))
        if lat is not None and "montserrat" in lat.get("typeface", "").lower():
            continue
        set_font(rpr, BOLD_FONT if rpr.get("b") == "1" else BODY_FONT)
        changed = True
    for tag in ("endParaRPr", "defRPr"):
        for rpr in el.iter(qn("a:" + tag)):
            lat = rpr.find(qn("a:latin"))
            if lat is not None and not lat.get("typeface", "").startswith("+") \
                    and "montserrat" not in lat.get("typeface", "").lower():
                set_font(rpr, BODY_FONT)
    return changed


def scale_text(el, f):
    """Уменьшает текст пропорционально f: fontScale в текстовых блоках, явные размеры в таблицах."""
    if abs(f - 1) < 0.01:
        return
    in_tbl = set()
    for tc in el.iter(qn("a:tc")):
        for node in tc.iter():
            in_tbl.add(node)
        for r in tc.iter(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                rpr = etree.Element(qn("a:rPr")); r.insert(0, rpr)
            sz = int(rpr.get("sz") or 1800)
            rpr.set("sz", str(max(int(sz * f), min(sz, 800))))
        for epr in tc.iter(qn("a:endParaRPr")):
            if epr.get("sz"):
                epr.set("sz", str(max(int(int(epr.get("sz")) * f), 600)))
    for bp in el.iter(qn("a:bodyPr")):
        if bp in in_tbl or bp.getparent().getparent() is None:
            continue
        cur = 1.0
        na = bp.find(qn("a:normAutofit"))
        if na is not None and na.get("fontScale"):
            cur = int(na.get("fontScale")) / 100000
        for tag in ("noAutofit", "normAutofit", "spAutoFit"):
            for ch in bp.findall(qn("a:" + tag)):
                bp.remove(ch)
        na = etree.Element(qn("a:normAutofit"))
        na.set("fontScale", str(max(25000, int(cur * f * 100000))))
        bp.insert(0, na)


def scale_tables(el, f):
    if abs(f - 1) < 0.01:
        return
    for c in el.iter(qn("a:gridCol")):
        c.set("w", str(int(int(c.get("w")) * f)))
    for r in el.iter(qn("a:tr")):
        r.set("h", str(int(int(r.get("h")) * f)))



# ================================================================ оценка размещения текста
def style_chain(slide, el):
    """Цепочка стилей текста: сам объект -> макет -> мастер (для поиска унаследованных размеров)."""
    chain = [el.find("./p:txBody/a:lstStyle", NS)]
    ph = el.find("./p:nvSpPr/p:nvPr/p:ph", NS)
    try:
        layout = slide.slide_layout
        master = layout.slide_master
        typ = ph.get("type") if ph is not None else None
        if ph is not None:
            for lph in layout.placeholders:
                if str(lph.placeholder_format.idx) == ph.get("idx", "0"):
                    chain.append(lph._element.find("./p:txBody/a:lstStyle", NS))
                    break
        name = "otherStyle" if ph is None else ("titleStyle" if typ in ("title", "ctrTitle") else "bodyStyle")
        chain.append(master._element.find(".//p:txStyles/p:" + name, NS))
    except Exception:
        pass
    return chain


def _lvl(chain, lvl, sub):
    for ch in chain:
        if ch is None:
            continue
        node = ch.find(qn(f"a:lvl{lvl}pPr"))
        if node is None:
            continue
        found = node.find(sub, NS) if sub else node
        if found is not None:
            return found
    return None


def est_text_height(c, chain, width_emu, comp, cw):
    """Грубая оценка высоты текста блока (EMU) с учётом шрифта, отступов, интервалов."""
    body = c.find(qn("p:txBody"))
    if body is None:
        return 0
    bp = body.find(qn("a:bodyPr"))
    li = int(bp.get("lIns", 91440)); ri = int(bp.get("rIns", 91440))
    ti = int(bp.get("tIns", 45720)); bi = int(bp.get("bIns", 45720))
    cur = 1.0
    na = bp.find(qn("a:normAutofit"))
    if na is not None and na.get("fontScale"):
        cur = int(na.get("fontScale")) / 100000
    total = 0.0
    for p in body.findall(qn("a:p")):
        ppr = p.find(qn("a:pPr"))
        lvl = int(ppr.get("lvl", 0)) + 1 if ppr is not None else 1
        # размер шрифта
        szs = [int(r.get("sz")) for r in p.iter(qn("a:rPr")) if r.get("sz")]
        if szs:
            sz = max(szs) / 100
        else:
            d = _lvl(chain, lvl, "a:defRPr")
            ep = p.find(qn("a:endParaRPr"))
            sz = int(ep.get("sz")) / 100 if ep is not None and ep.get("sz") else \
                (int(d.get("sz")) / 100 if d is not None and d.get("sz") else 18)
        sz *= cur * comp
        # отступ слева
        marl = None
        if ppr is not None and ppr.get("marL") is not None:
            marl = int(ppr.get("marL"))
        else:
            lv = _lvl(chain, lvl, None)
            marl = int(lv.get("marL", 0)) if lv is not None else 0
        # межстрочный и интервалы
        lns = 1.0
        for src in (ppr, _lvl(chain, lvl, None)):
            if src is not None:
                sp = src.find("./a:lnSpc/a:spcPct", NS)
                if sp is not None:
                    lns = int(sp.get("val")) / 100000
                    break
        def space(tag):
            for src in (ppr, _lvl(chain, lvl, None)):
                if src is not None:
                    pc = src.find(f"./a:{tag}/a:spcPct", NS)
                    if pc is not None:
                        return int(pc.get("val")) / 100000 * sz
                    pt = src.find(f"./a:{tag}/a:spcPts", NS)
                    if pt is not None:
                        return int(pt.get("val")) / 100
            return 0
        # текст по строкам (a:br - принудительный перенос)
        segs, cur_t = [], ""
        for ch in p:
            if ch.tag == qn("a:br"):
                segs.append(cur_t); cur_t = ""
            elif ch.tag in (qn("a:r"), qn("a:fld")):
                cur_t += "".join(t.text or "" for t in ch.iter(qn("a:t")))
        segs.append(cur_t)
        avail = max(width_emu - li - ri - marl, 300000)
        lines = sum(max(1, n_lines(t, avail, sz, cw)) if t.strip() else 1 for t in segs)
        total += lines * sz * 1.2 * lns + space("spcBef") + space("spcAft")
    return int(total * 12700 + ti + bi)


def is_visible(c):
    sp = c.find(qn("p:spPr"))
    if sp is None:
        return False
    fill = any(sp.find(qn("a:" + t)) is not None for t in ("solidFill", "gradFill", "blipFill", "pattFill"))
    ln = sp.find(qn("a:ln"))
    has_ln = ln is not None and ln.find(qn("a:noFill")) is None and \
        any(ln.find(qn("a:" + t)) is not None for t in ("solidFill", "gradFill", "pattFill"))
    styled = c.find(qn("p:style")) is not None and not (
        sp.find(qn("a:noFill")) is not None and (ln is not None and ln.find(qn("a:noFill")) is not None))
    return fill or has_ln or styled


def anchor_top(c):
    bp = c.find("./p:txBody/a:bodyPr", NS)
    return bp is None or bp.get("anchor", "t") == "t"


# ================================================================ рамки шаблона
def find_protos(prs, ov):
    slides = list(prs.slides)

    def texts(s):
        return [sp_text(e) for e in s.shapes._spTree if e.tag == qn("p:sp")]

    def has_bg(s):
        return s._element.find(".//p:bg", NS) is not None

    pick = {}
    for i, s in enumerate(slides):
        t = " ".join(texts(s))
        if has_bg(s) and "Название" in t and "title" not in pick: pick["title"] = i
        if has_bg(s) and "раздела" in t and "section" not in pick: pick["section"] = i
        has_line = s._element.find(".//p:cxnSp", NS) is not None
        if has_line and "Презентация" in t and "content" not in pick: pick["content"] = i
    for k, v in (ov or {}).items():
        if v: pick[k] = v - 1
    missing = [k for k in ("title", "section", "content") if k not in pick]
    if missing:
        raise SystemExit(f"Не нашёл в шаблоне рамки: {missing}. Укажите номера слайдов: "
                         f"--proto-title N --proto-section N --proto-content N")
    return {k: slides[v] for k, v in pick.items()}, {k: v + 1 for k, v in pick.items()}


def proto_roles_content(proto):
    roles = {}
    for el in proto.shapes._spTree:
        if el.tag == qn("p:cxnSp"):
            roles["line"] = el
        elif el.tag == qn("p:sp") and sp_text(el).strip():
            x, y, cx, cy = get_xfrm(el); t = sp_text(el).strip()
            if y < 700000 and x < 1500000: roles["label"] = el
            elif y < 700000 and t.isdigit() and x > 9000000: roles["number"] = el
            elif 900000 < y < 1600000 and x < 1300000 and "title" not in roles: roles["title"] = el
    return roles


def clone_frame(proto, new_slide, package, deferred, shapes_filter=None):
    """Копирует фон и объекты рамки в новый слайд. Возвращает список клонов в порядке оригиналов."""
    bg = proto._element.find(".//p:bg", NS)
    if bg is not None:
        cs = new_slide._element.find(qn("p:cSld"))
        old = cs.find(qn("p:bg"))
        if old is not None: cs.remove(old)
        cs.insert(0, copy.deepcopy(bg))
    rid_map = copy_rels(proto.part, new_slide.part, package, {}, deferred,
                        skip_types=(RT.SLIDE_LAYOUT, RT.NOTES_SLIDE), same_package=True)
    out = []
    tree = new_slide.shapes._spTree
    used = [int(i) for i in new_slide._element.xpath("//p:cNvPr/@id")]
    nxt = max(used + [1]) + 1
    for el in proto.shapes._spTree:
        if el.tag in SKIP_TAGS or (shapes_filter and not shapes_filter(el)):
            continue
        c = copy.deepcopy(el)
        remap_rids(c, rid_map)
        for cnv in c.iter(qn("p:cNvPr")):
            cnv.set("id", str(nxt)); nxt += 1
        tree.insert_element_before(c, "p:extLst")
        out.append((el, c))
    return out


# ================================================================ определение заголовка
def is_title(sh):
    return sh.is_placeholder and sh.placeholder_format.type in TITLE_TYPES


def find_title(slide, slide_h, no_guess):
    for sh in slide.shapes:
        if is_title(sh) and sh.has_text_frame and sh.text_frame.text.strip():
            return sh, False
    for sh in slide.shapes:   # текстовый блок, названный "Заголовок"/"Title"
        if sh.has_text_frame and sh.text_frame.text.strip() and re.search(r"заголовок|title", sh.name, re.I) \
                and sh.top is not None and sh.top < slide_h * 0.3 and len(sh.text_frame.text) <= 200:
            return sh, False
    if no_guess:
        return None, False
    c = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip() and s.top is not None
         and s.top < slide_h * 0.2 and len(s.text_frame.text) <= 120]
    return (min(c, key=lambda s: (s.top, s.left or 0)), True) if c else (None, False)


def geom(sh):
    try:
        if None in (sh.left, sh.top, sh.width, sh.height):
            return None
        return sh.left, sh.top, sh.width, sh.height
    except Exception:
        return None


# ================================================================ проверка
def words(els):
    c = Counter()
    for e in els:
        c.update(re.findall(r"\w+", " ".join(t.text or "" for t in e.iter(qn("a:t"))).lower()))
    return c


def stats(els):
    s = Counter()
    for e in els:
        for n in e.iter():
            if n.tag == qn("p:pic"): s["картинок"] += 1
            elif n.tag == qn("a:tbl"): s["таблиц"] += 1
            elif n.tag == qn("p:grpSp"): s["групп"] += 1
            elif n.tag in (qn("a:videoFile"), qn("a:audioFile")): s["медиа"] += 1
            elif n.tag == qn("a:graphicData") and n.get("uri", "").endswith("/chart"): s["диаграмм"] += 1
    return s


def verify(src_els, dst_els, dst_slide, n):
    out = []
    a, b = stats(src_els), stats(dst_els)
    for k in set(a) | set(b):
        if a[k] != b[k]:
            out.append(f"слайд {n}: {k} было {a[k]}, стало {b[k]}")
    miss = words(src_els) - words([dst_slide._element])
    if miss:
        out.append(f"слайд {n}: пропали слова: {', '.join(list(miss)[:12])}")
    return out


# ================================================================ основная логика
def convert_one(template, src_path, out_path, o):
    src = Presentation(str(src_path))
    prs = Presentation(str(template))
    sw, sh_ = prs.slide_width, prs.slide_height
    ow, oh = src.slide_width, src.slide_height
    base = min(sw / ow, sh_ / oh)
    protos, _ = find_protos(prs, o.get("proto"))
    orig_ids = list(prs.slides._sldIdLst)
    layout_obj = next((l for l in prs.slide_layouts if "объект" in l.name.lower() and "заголовок" in l.name.lower()
                       and "два" not in l.name.lower() and "верт" not in l.name.lower()), prs.slide_layouts[1])
    layout_idx1 = any(p.placeholder_format.idx == 1 for p in layout_obj.placeholders)
    package, memo, deferred, warnings = prs.part.package, {}, [], []
    slide_map, pairs = {}, []
    label, counter = o.get("label") or "Презентация", 0

    for n, s in enumerate(src.slides, 1):
        title_sh, guessed = find_title(s, oh, o.get("no_guess"))
        if title_sh is None and n == 1:
            def _sz(sh):
                v = [int(x) for x in sh._element.xpath(".//a:rPr/@sz") if x.isdigit()]
                return max(v) if v else 0
            cands = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.strip()]
            if cands:
                title_sh = max(cands, key=lambda x: (_sz(x), -(x.top or 0)))
                guessed = True
        if guessed:
            warnings.append(f"слайд {n}: заголовок определён по размеру/положению (не placeholder)")
        tree = s.shapes._spTree
        content_els, empty_ph = [], 0
        for el in tree:
            if el.tag in SKIP_TAGS or (title_sh is not None and el is title_sh._element):
                continue
            ph = el.find("./p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is not None:
                w = SlideShapeFactory(el, s.shapes)
                if w.placeholder_format.type in DROP_PH:
                    continue
                if not sp_text(el).strip():       # пустой плейсхолдер - только подсказка в редакторе
                    empty_ph += 1; continue
            content_els.append(el)
        src_all = [e for e in tree if e.tag not in SKIP_TAGS]
        title_lines = [] if title_sh is None else [paras(title_sh._element)] and paras(title_sh._element)
        title_str = " ".join(title_lines)
        text_only = all(e.tag == qn("p:sp") and sp_text(e).strip() for e in content_els)

        if n == 1 and title_sh is not None:
            kind = "title"
        elif title_sh is not None and not content_els:
            kind = "section"
        else:
            kind = "content"

        new = prs.slides.add_slide(layout_obj)
        for ph in list(new.placeholders):
            ph._element.getparent().remove(ph._element)
        slide_map[s.part] = new.part
        dst_els = []

        # ---------- титульный слайд
        if kind == "title":
            cl = clone_frame(protos["title"], new, package, deferred)
            tb = sorted([(get_xfrm(c)[1], c) for _, c in cl if c.tag == qn("p:sp") and sp_text(c).strip()],
                        key=lambda t: t[0])
            t_el, st_el = tb[0][1], (tb[1][1] if len(tb) > 2 else None)
            x, y, cx, cy = get_xfrm(t_el)
            nxt_y = get_xfrm(tb[1][1])[1] if len(tb) > 1 else y + cy
            bsz, lns = run_props(t_el)
            pt, ln, h = fit_pt(title_str, cx, nxt_y - y - 200000, bsz, lns, 0.64, 22)
            fill_shape(t_el, [title_str], pt); set_xfrm(t_el, cy=h + 50000)
            extra = []
            for e in sorted(content_els, key=lambda e: (get_xfrm(e)[1], get_xfrm(e)[0]) if e.find(".//a:xfrm", NS) is not None else (0, 0)):
                extra += paras(e)
            if st_el is not None:
                if extra:
                    sx, sy, scx, scy = get_xfrm(st_el)
                    date_y = get_xfrm(tb[-1][1])[1]
                    bsz2, lns2 = run_props(st_el)
                    stext = " ".join(extra)
                    pt2, ln2, h2 = fit_pt(stext, min(scx, 5500000), date_y - sy - 250000, bsz2, 1.0, 0.62, 9)
                    if h2 > date_y - sy - 250000:
                        warnings.append(f"слайд {n}: текст подзаголовка очень длинный, проверьте глазами")
                    fill_shape(st_el, extra, pt2); set_xfrm(st_el, cx=min(scx, 5500000), cy=date_y - sy - 250000)
                else:
                    st_el.getparent().remove(st_el)
            dst_els = [t_el] + ([st_el] if extra else [])
            dropped = stats([e for e in content_els if not (e.tag == qn("p:sp") and sp_text(e).strip())])
            if dropped:
                warnings.append("слайд 1: не перенесены нетекстовые объекты старого титула ("
                                + ", ".join(f"{k}: {v}" for k, v in dropped.items()) + ")")
            # исходные тексты титульного слайда учтены в подзаголовке; для проверки сравниваем слова
            dst_check = [new._element]
        # ---------- раздел
        elif kind == "section":
            cl = clone_frame(protos["section"], new, package, deferred)
            t_el = next(c for _, c in cl if c.tag == qn("p:sp") and sp_text(c).strip())
            x, y, cx, cy = get_xfrm(t_el)
            bsz, lns = run_props(t_el)
            pt, ln, h = fit_pt(title_str, cx, 2400000, bsz, lns, 0.64, 26)
            fill_shape(t_el, [title_str], pt); set_xfrm(t_el, cy=h + 50000)
            label = title_str[:24]
            dst_els = [t_el]; dst_check = [new._element]
        # ---------- обычный слайд
        else:
            counter += 1
            roles = proto_roles_content(protos["content"])
            keep = lambda el: el in roles.values()
            cl = clone_frame(protos["content"], new, package, deferred, shapes_filter=keep)
            by_role = {r: c for r, el in roles.items() for o_, c in cl if o_ is el}
            lab = label[:30]
            fill_shape(by_role["label"], [lab])
            set_xfrm(by_role["label"], cx=len(lab) * 140000 + 100000)
            lx = 657696 + len(lab) * 133000 + 430000
            set_xfrm(by_role["line"], x=lx, cx=max(LINE_END - lx, 500000))
            fill_shape(by_role["number"], [f"{counter:02d}"])
            top = TITLE_TOP
            if title_sh is not None:
                t_el = by_role["title"]
                bsz, lns = run_props(t_el)
                bsz = min(bsz, o.get("title_pt", 32))
                width = RIGHT - LEFT
                pt, ln, h = fit_pt(title_str, width, 10 ** 9, bsz, lns, 0.60, 18)
                while ln > 2 and pt > 18:
                    pt -= 1; ln = n_lines(title_str, width, pt, 0.60)
                h = ln * pt * 1.2 * lns / 72 * 914400
                fill_shape(t_el, [title_str], pt)
                set_xfrm(t_el, x=LEFT, y=TITLE_TOP, cx=width, cy=h + 30000)
                top = TITLE_TOP + h + GAP
                dst_els.append(t_el)
            else:
                t_el = by_role["title"]; t_el.getparent().remove(t_el)
                warnings.append(f"слайд {n}: заголовок не найден")
            dst_check = [new._element]

            # --- содержимое
            rid_map = copy_rels(s.part, new.part, package, memo, deferred,
                                skip_types=(RT.SLIDE_LAYOUT, RT.NOTES_SLIDE))
            id_map = {}
            used = [int(i) for i in new._element.xpath("//p:cNvPr/@id")]
            nid = max(used + [1]) + 1
            kept_obj = False
            fcomp = o.get("font_comp", 1.0)
            items = []
            for el in content_els:
                chain = style_chain(s, el)
                c = copy.deepcopy(el)
                remap_rids(c, rid_map)
                for cnv in c.iter(qn("p:cNvPr")):
                    id_map[cnv.get("id")] = str(nid); cnv.set("id", str(nid)); nid += 1
                ph = c.find("./p:nvSpPr/p:nvPr/p:ph", NS)
                if ph is not None:
                    typ, idx = ph.get("type") or "obj", ph.get("idx", "0")
                    if layout_idx1 and typ in ("obj", "body") and idx == "1" and not kept_obj:
                        kept_obj = True
                    else:
                        ph.getparent().remove(ph)
                restyled = (not o.get("keep_fonts")) and restyle_fonts(c)
                comp = fcomp if restyled else 1.0
                cw = 0.57 if restyled else 0.5
                try:
                    g = geom(SlideShapeFactory(el, s.shapes))
                except Exception:
                    g = None
                it = {"c": c, "g": g, "comp": comp, "cw": cw, "chain": chain, "need": None,
                      "visible": True, "top": True, "s": 1.0, "h": g[3] if g else None}
                if g and c.tag == qn("p:sp") and sp_text(c).strip():
                    it["visible"], it["top"] = is_visible(c), anchor_top(c)
                    it["need"] = est_text_height(c, chain, g[2], comp, cw)
                    if it["need"] > g[3]:            # текст не помещается в свой блок
                        it["s"] = max(0.5, (g[3] / it["need"]) ** 0.5 * 0.97)
                    elif not it["visible"] and it["top"]:
                        it["h"] = max(int(it["need"] * 1.05), 200000)   # без рамки/заливки блок можно ужать до текста
                items.append(it)

            bl, bt, bw = LEFT, top, RIGHT - LEFT
            bh = sh_ - BOTTOM_MARGIN - bt
            single = len(items) == 1 and items[0]["need"] is not None and not items[0]["visible"] \
                and items[0]["g"] is not None
            if single:
                # единственный текстовый блок: используем всю ширину сетки (меньше строк - крупнее текст)
                it = items[0]
                need = est_text_height(it["c"], it["chain"], bw, it["comp"], it["cw"])
                sc = 1.0 if need <= bh else max(0.45, (bh / need) ** 0.5 * 0.97)
                it["s"], it["place"] = sc, (bl, bt, bw, int(min(need * 1.05, bh)) if sc == 1.0 else bh)
                f = 1.0
            else:
                gs = [(it["g"][0], it["g"][1], it["g"][2], it["h"]) for it in items if it["g"]]
                if gs:
                    x0 = min(g[0] for g in gs); y0 = min(g[1] for g in gs)
                    x1 = max(g[0] + g[2] for g in gs); y1 = max(g[1] + g[3] for g in gs)
                    f = min(base, bw / max(x1 - x0, 1), bh / max(y1 - y0, 1))
                else:
                    f, x0, y0 = base, 0, 0
                for it in items:
                    if it["g"]:
                        g = it["g"]
                        it["place"] = (int(bl + (g[0] - x0) * f), int(bt + (it["g"][1] - y0) * f),
                                       int(g[2] * f), int(it["h"] * f))
            for it in items:
                c = it["c"]
                scale_text(c, (f if not single else 1.0) * it["comp"] * it["s"])
                scale_tables(c, f)
                new.shapes._spTree.insert_element_before(c, "p:extLst")
                dst_els.append(c)
                if it["g"]:
                    try:
                        nw = SlideShapeFactory(c, new.shapes)
                        x, y, w_, h_ = it["place"]
                        nw.left, nw.top, nw.width, nw.height = int(x), int(y), int(w_), int(h_)
                    except Exception as e:
                        warnings.append(f"слайд {n}: не удалось пересчитать положение объекта: {e}")
                else:
                    warnings.append(f"слайд {n}: объект без явного положения оставлен как есть")
            for tag in ("stCxn", "endCxn"):
                for cx_ in new._element.iter(qn("a:" + tag)):
                    if cx_.get("id") in id_map: cx_.set("id", id_map[cx_.get("id")])
            timing = s._element.find("p:timing", NS)
            if timing is not None:
                t = copy.deepcopy(timing)
                spids = t.xpath(".//@spid")
                if all(x in id_map for x in spids):
                    for node in t.iter():
                        if node.get("spid") in id_map: node.set("spid", id_map[node.get("spid")])
                    new._element.append(t)
                else:
                    warnings.append(f"слайд {n}: анимация не перенесена (привязана к заголовку/удалённым объектам)")
            if empty_ph:
                warnings.append(f"слайд {n}: пропущено пустых плейсхолдеров: {empty_ph} (в них не было содержимого)")

        # заметки докладчика
        if s.has_notes_slide and s.notes_slide.notes_text_frame is not None:
            txt = s.notes_slide.notes_text_frame.text
            if txt.strip():
                new.notes_slide.notes_text_frame.text = txt
        pairs.append((src_all, dst_els, new, kind))

    # переходы-гиперссылки между слайдами
    for dst_part, rid, tgt in deferred:
        nt = slide_map.get(tgt)
        if nt is None:
            warnings.append("ссылка на слайд вне презентации пропущена"); continue
        new_rid = dst_part.relate_to(nt, RT.SLIDE)
        for el in dst_part._element.iter():
            for k, v in list(el.attrib.items()):
                if k.startswith("{" + R_NS + "}") and v == rid:
                    el.set(k, new_rid)

    # убираем исходные слайды шаблона
    lst = prs.slides._sldIdLst
    for sid in orig_ids:
        prs.part.drop_rel(sid.rId); lst.remove(sid)

    problems = []
    for i, (src_all, dst_els, new, kind) in enumerate(pairs, 1):
        if kind == "content":
            problems += verify(src_all, dst_els, new, i)
        else:   # рамки с декором: сравниваем только слова и неделимые элементы
            miss = words(src_all) - words([new._element])
            if miss:
                problems.append(f"слайд {i}: пропали слова: {', '.join(list(miss)[:12])}")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    Presentation(str(out_path))
    return len(pairs), warnings, problems


def inspect(template, ov):
    prs = Presentation(template)
    print(f"Размер слайда: {Emu(prs.slide_width).inches:.2f} x {Emu(prs.slide_height).inches:.2f} дюйма")
    protos, nums = find_protos(prs, ov)
    print("Рамки шаблона (номера слайдов):", nums)
    print("Роли в рамке со шапкой:", {k: sp_text(v)[:20] for k, v in proto_roles_content(protos["content"]).items()})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--src"); ap.add_argument("--out")
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--label", help="надпись слева в шапке (по умолчанию 'Презентация'; после слайда-раздела берётся его название)")
    ap.add_argument("--proto-title", type=int); ap.add_argument("--proto-section", type=int)
    ap.add_argument("--proto-content", type=int)
    ap.add_argument("--keep-fonts", action="store_true", help="не менять шрифты исходника на Montserrat")
    ap.add_argument("--font-comp", type=float, default=1.0,
                    help="дополнительное уменьшение текста при смене шрифта (1 = нет; текст и так уменьшается только там, где не помещается)")
    ap.add_argument("--no-guess", action="store_true")
    ap.add_argument("--title-pt", type=int, default=32, help="максимальный размер заголовка на обычных слайдах (в шаблоне 40)")
    a = ap.parse_args()
    ov = {"title": a.proto_title, "section": a.proto_section, "content": a.proto_content}
    if a.inspect:
        return inspect(a.template, ov)
    if not (a.src and a.out):
        ap.error("нужны --src и --out")
    opts = {"proto": ov, "label": a.label, "keep_fonts": a.keep_fonts,
            "font_comp": a.font_comp, "no_guess": a.no_guess, "title_pt": a.title_pt}
    src_dir, out_dir = Path(a.src), Path(a.out)
    files = sorted(p for p in src_dir.rglob("*.pptx") if not p.name.startswith("~$"))
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    for i, f in enumerate(files, 1):
        dst = out_dir / f.relative_to(src_dir)
        try:
            n, warns, probs = convert_one(a.template, f, dst, opts)
            rows.append([f.name, "ПРОВЕРИТЬ" if probs else "ok", n, " | ".join(probs), " | ".join(warns)])
        except Exception as e:
            rows.append([f.name, "ОШИБКА", 0, f"{type(e).__name__}: {e}", ""])
            traceback.print_exc()
        print(f"[{i}/{len(files)}] {f.name}: {rows[-1][1]}")
    with open(out_dir / "report.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["файл", "статус", "слайдов", "потери/расхождения", "предупреждения"])
        w.writerows(rows)
    print(f"\nГотово. Отчёт: {out_dir / 'report.csv'}")


if __name__ == "__main__":
    main()
