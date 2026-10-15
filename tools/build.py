"""Construction du Guide du marché CT à partir de content.yaml (rejouable).

    python3 tools/build.py [--content livrables/content.yaml] [--out livrables/Guide_du_marche_CT_2026.pptx]

- Part du .pptx de couverture (masque + thème « Guide du Marché CT Philips ») : toutes les
  slides utilisent ses mises en page ; la slide de couverture d'origine est conservée.
- Tout le texte vient de content.yaml ; ce script ne contient aucune donnée produit.
- Couleurs : uniquement la palette du thème (+ teintes calculées par mélange de deux couleurs du thème).
- Les marqueurs [À COMPLÉTER] / [À VÉRIFIER] / [VISUEL À AJOUTER] sont stylés à part et
  recensés automatiquement dans A_COMPLETER.md ; les sources vont dans les notes.
"""
import argparse, copy, re, sys
from pathlib import Path

import yaml
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TODO, CHECK, VISUAL = "[À COMPLÉTER]", "[À VÉRIFIER]", "[VISUEL À AJOUTER]"
FONT = {False: "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        True: "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"}
SW, SH = 13.333, 7.5
MX = 0.6  # marge latérale


# --------------------------------------------------------------------------- palette
class Palette:
    def __init__(self, theme_xml):
        self.c = dict(re.findall(r"<a:(dk1|lt1|dk2|lt2|accent\d)><a:srgbClr val=\"(\w{6})\"", theme_xml))

    def __getitem__(self, key):
        """'accent1', 'dk1', ou mélange 'accent1/dk1/0.4' (40 % de la 2e couleur)."""
        if "/" in key:
            a, b, t = key.split("/")
            ca, cb, t = self.rgb(a), self.rgb(b), float(t)
            return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))
        return self.c[key]

    def rgb(self, key):
        h = self[key]
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

    def lum(self, key):
        def ch(v):
            v /= 255
            return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
        r, g, b = (ch(v) for v in self.rgb(key))
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    def on(self, key):
        """Couleur de texte lisible sur un fond de la palette."""
        return "dk1" if self.lum(key) > 0.25 else "lt1"


# --------------------------------------------------------------------------- texte
_fonts = {}


def text_width(s, size, bold=False):
    k = (size, bold)
    if k not in _fonts:
        _fonts[k] = ImageFont.truetype(FONT[bold], int(size * 10))
    return _fonts[k].getlength(s) / 10 / 72  # pouces


def wrap_lines(text, size, width, bold=False):
    n = 0
    for para in str(text).split("\n"):
        line = ""
        for w in para.split(" "):
            t = (line + " " + w).strip()
            if text_width(t, size, bold) <= width or not line:
                line = t
            else:
                n += 1
                line = w
        n += 1
    return n


class Builder:
    def __init__(self, content_path, out_path):
        self.content_path = Path(content_path)
        self.content = yaml.safe_load(self.content_path.read_text())
        meta = self.content["meta"]
        self.out_path = Path(out_path)
        self.img_root = ROOT / meta["images_root"]
        self.prs = Presentation(ROOT / meta["base_pptx"])
        theme = self.prs.slide_master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme").blob.decode()
        self.P = Palette(theme)
        self.meta = meta
        self.layout_blank = self.prs.slide_layouts.get_by_name("DEFAULT")
        self.layout_title = self.prs.slide_layouts.get_by_name("Title and body")
        self.ids, self.links, self.todos, self.warnings = {}, [], [], []
        self.slide_no = 0
        self._img_index = {}
        self.section = None

    # ------------------------------------------------------------------ primitives
    def color(self, key):
        return RGBColor.from_string(self.P[key])

    def rect(self, s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None,
             dash=False, line_w=1.0, name=None):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            sh.name = name
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = self.color(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = self.color(line)
            sh.line.width = Pt(line_w)
            if dash:
                sh.line.dash_style = 4  # dash
        else:
            sh.line.fill.background()
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = radius
        sh.shadow.inherit = False
        tf = sh.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        return sh

    def gradient(self, sh, stops, angle=0):
        """stops = [(pos 0-1, couleur, alpha 0-1)] ; dégradé linéaire natif."""
        spPr = sh._element.spPr
        for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
            for el in spPr.findall(qn(tag)):
                spPr.remove(el)
        g = etree.SubElement(spPr, qn("a:gradFill"), rotWithShape="1")
        lst = etree.SubElement(g, qn("a:gsLst"))
        for pos, col, alpha in stops:
            gs = etree.SubElement(lst, qn("a:gs"), pos=str(int(pos * 100000)))
            c = etree.SubElement(gs, qn("a:srgbClr"), val=self.P[col])
            if alpha < 1:
                etree.SubElement(c, qn("a:alpha"), val=str(int(alpha * 100000)))
        etree.SubElement(g, qn("a:lin"), ang=str(int(angle * 60000)), scaled="0")
        # gradFill doit précéder a:ln
        ln = spPr.find(qn("a:ln"))
        if ln is not None:
            spPr.remove(g)
            ln.addprevious(g)

    def text(self, s, x, y, w, h, runs, size=16, color="lt1", bold=False, align="l", anchor="t",
             italic=False, spacing=None, line=1.0, name=None, check=True, para_space=0):
        """runs : str, ou liste de paragraphes ; un paragraphe = str ou liste de (texte, {opts})."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        paras = runs if isinstance(runs, list) else [runs]
        plain = []
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            p.line_spacing = line
            if para_space and i:
                p.space_before = Pt(para_space)
            segs = para if isinstance(para, list) else [(para, {})]
            ptxt = ""
            for seg in segs:
                t, o = (seg, {}) if isinstance(seg, str) else seg
                for part in self.split_markers(str(t)):
                    r = p.add_run()
                    r.text = part
                    f = r.font
                    f.size = Pt(o.get("size", size))
                    f.bold = o.get("bold", bold)
                    f.italic = o.get("italic", italic)
                    col = o.get("color", color)
                    if part in (TODO, CHECK, VISUAL):
                        f.italic, f.bold = True, True
                        col = "accent4" if part == TODO else "accent1"
                        f.size = Pt(min(o.get("size", size), 14))
                    f.color.rgb = self.color(col)
                    if spacing:
                        r.font._element.set("spc", str(spacing))
                    ptxt += part
            plain.append((ptxt, max([size] + [o.get("size", size) for _, o in
                                              [(sg, {}) if isinstance(sg, str) else sg for sg in segs]]),
                          bold))
        if check:
            need = sum(wrap_lines(t, sz, w, b) * sz * 1.2 * line / 72 + (para_space / 72 if para_space else 0)
                       for t, sz, b in plain)
            if need > h + 0.05:
                self.warnings.append(f"slide {self.slide_no}: texte trop long ({need:.2f} > {h:.2f} in) : "
                                     f"{plain[0][0][:50]!r}")
        return tb

    @staticmethod
    def split_markers(t):
        return [p for p in re.split(r"(\[À COMPLÉTER\]|\[À VÉRIFIER\]|\[VISUEL À AJOUTER\])", t) if p]

    def image(self, s, ref, x, y, w, h, mode="cover", name=None, focus=(0.5, 0.5)):
        """cover : remplit la boîte en recadrant ; contain : tient dans la boîte. Jamais d'étirement."""
        path = self.resolve_image(ref)
        if path is None or not path.exists():
            self.warnings.append(f"slide {self.slide_no}: image absente {ref}")
            return self.visual_placeholder(s, x, y, w, h, ref)
        with Image.open(path) as im:
            iw, ih = im.size
        ir, br = iw / ih, w / h
        if mode == "contain":
            if ir > br:
                nw, nh = w, w / ir
            else:
                nw, nh = h * ir, h
            pic = s.shapes.add_picture(str(path), Inches(x + (w - nw) / 2), Inches(y + (h - nh) / 2),
                                       Inches(nw), Inches(nh))
            dpi = iw / nw
        else:
            pic = s.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
            if ir > br:  # trop large : recadrer gauche/droite
                keep = br / ir
                cut = 1 - keep
                pic.crop_left = cut * focus[0]
                pic.crop_right = cut * (1 - focus[0])
            else:
                keep = ir / br
                cut = 1 - keep
                pic.crop_top = cut * focus[1]
                pic.crop_bottom = cut * (1 - focus[1])
            dpi = (iw * (keep if ir > br else 1)) / w
        if dpi < 90:
            self.warnings.append(f"slide {self.slide_no}: résolution faible {ref} ({dpi:.0f} px/in)")
            self.todos.append((self.slide_no, self.cur_title,
                               f"Visuel {ref} : meilleure résolution souhaitable ({dpi:.0f} px/pouce à l'affichage ; "
                               f"c'est la meilleure version présente dans les sources)"))
        if name:
            pic.name = name
        pic.shadow.inherit = False
        return pic

    def resolve_image(self, ref):
        """« J-47 » = plus grande image de la page 47 du JFR ; « J-47#2 » = 2e plus grande ;
        sinon chemin relatif à analyse/."""
        m = re.fullmatch(r"([GJ])-(\d+)(?:#(\d+))?", str(ref))
        if not m:
            return self.img_root / ref
        doc = {"G": "guide_2025", "J": "jfr_2026"}[m.group(1)]
        if doc not in self._img_index:
            import json
            self._img_index[doc] = json.loads((self.img_root / f"{doc}_images.json").read_text())
        page, rank = int(m.group(2)), int(m.group(3) or 1)
        cands = sorted((i for i in self._img_index[doc] if page in i["pages"]), key=lambda i: -i["w"] * i["h"])
        return self.img_root / cands[rank - 1]["file"] if len(cands) >= rank else None

    @staticmethod
    def round_pic(pic, radius):
        """Cadre arrondi natif (prstGeom roundRect) sur une image."""
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        geom.set("prst", "roundRect")
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = etree.SubElement(geom, qn("a:avLst"))
        for g in list(av):
            av.remove(g)
        etree.SubElement(av, qn("a:gd"), name="adj", fmla=f"val {int(radius * 100000)}")

    def has_alpha(self, ref):
        p = self.resolve_image(ref)
        with Image.open(p) as im:
            return im.mode in ("RGBA", "LA", "P") and im.getextrema()[-1][0] < 250

    def visual_placeholder(self, s, x, y, w, h, hint=""):
        self.rect(s, x, y, w, h, fill="dk2", line="accent4", dash=True, name="Visuel à ajouter")
        if w < 1.6:
            self.text(s, x + 0.05, y, w - 0.1, h, [[("VISUEL", {"size": 9})], [("À AJOUTER", {"size": 9})]],
                      color="accent4", bold=True, italic=True, align="c", anchor="m", check=False)
        else:
            self.text(s, x, y + h / 2 - 0.3, w, 0.6, [[(VISUAL, {"size": 14})],
                                                       [(hint, {"size": 10, "color": "lt2"})]],
                      align="c", anchor="m", check=False)
        self.todos.append((self.slide_no, self.cur_title, VISUAL + (f" – {hint}" if hint else "")))

    def todo_box(self, s, x, y, w, h, label="", size=12):
        """Emplacement de donnée manquante : cadre pointillé gris + texte italique."""
        self.rect(s, x, y, w, h, line="accent4", dash=True, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2,
                  name="À compléter")
        self.text(s, x, y, w, h, TODO, size=size, align="c", anchor="m", check=False)

    def badge_new(self, s, x, y):
        b = self.rect(s, x, y, 0.95, 0.3, fill="accent1", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5,
                      name="Badge Nouveau")
        self.text(s, x, y, 0.95, 0.3, "NOUVEAU", size=10, bold=True, color="dk1", align="c", anchor="m",
                  spacing=100, check=False)
        return b

    def token(self, s, x, y, d, label, fill="dk2", line="accent1", color="accent1", size=11):
        """Pictogramme homogène : pastille ronde + abréviation d'unité."""
        self.rect(s, x, y, d, d, fill=fill, line=line, shape=MSO_SHAPE.OVAL, line_w=1.25, name="Picto")
        self.text(s, x, y, d, d, label, size=size, bold=True, color=color, align="c", anchor="m", check=False)

    def link(self, shape, target):
        self.links.append((shape, target))

    # ------------------------------------------------------------------ cadre commun
    def new_slide(self, sd, layout="title"):
        s = self.prs.slides.add_slide(self.layout_title if layout == "title" else self.layout_blank)
        self.slide_no = len(self.prs.slides)
        self.cur_title = sd.get("title", "")
        for ph in list(s.placeholders):
            if ph.placeholder_format.type != 1:  # garder seulement le titre
                ph._element.getparent().remove(ph._element)
        if sd.get("id"):
            self.ids[sd["id"]] = s
        self.collect_markers(sd)
        self.notes(s, sd)
        return s

    def set_title(self, s, title, x=MX, y=0.78, w=SW - 2 * MX, h=0.75, size=30, color="lt1"):
        ph = s.shapes.title
        if ph is None:
            return self.text(s, x, y, w, h, title, size=size, bold=True, color=color)
        ph.left, ph.top, ph.width, ph.height = Inches(x), Inches(y), Inches(w), Inches(h)
        tf = ph.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = title
        r.font.size, r.font.bold = Pt(size), True
        r.font.color.rgb = self.color(color)
        lines = wrap_lines(title, size, w, True)
        self.title_lines = lines
        if lines * size * 1.15 / 72 > h + 0.05 and not getattr(self, "_allow_2l", False):
            self.warnings.append(f"slide {self.slide_no}: titre trop long {title!r}")
        return ph

    def chrome(self, s, sd, dark=True, show_section=True):
        """En-tête de section, pied de page, bouton sommaire, numéro, mention de confidentialité."""
        if s.shapes.title is not None:  # le titre (placeholder du masque) passe au premier plan
            el = s.shapes.title._element
            tree = el.getparent()
            tree.remove(el)
            tree.append(el)
        sec = self.section
        fg = "lt1" if dark else "dk1"
        if show_section and sec:
            label = sec["short"].upper()
            w = text_width(label, 9, True) + 0.5
            self.rect(s, MX, 0.36, 0.16, 0.16, fill=sec["color"], shape=MSO_SHAPE.OVAL, name="Code section")
            t = self.text(s, MX + 0.26, 0.31, w + 2, 0.26, label, size=9, bold=True,
                          color=sec["color"] if dark else "dk2", spacing=150, check=False)
            self.link(t, sec["id"])
        # bouton sommaire
        b = self.rect(s, MX, 6.92, 1.25, 0.32, line="lt2" if dark else "dk2", line_w=0.75,
                      shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5, name="Retour sommaire")
        self.text(s, MX, 6.92, 1.25, 0.32, "‹  SOMMAIRE", size=9, bold=True, color=fg, align="c", anchor="m",
                  spacing=80, check=False)
        self.link(b, "sommaire")
        conf = self.meta["confidentiality"] + "  ·  " + self.meta["copyright"]
        for extra in sd.get("confidentiality_extra", []):
            conf = extra + "  ·  " + conf
        self.text(s, 2.1, 6.98, 8.6, 0.22, conf, size=8, color="lt2" if dark else "dk2", check=False)
        self.text(s, SW - MX - 2.2, 6.95, 1.0, 0.25, str(self.slide_no), size=9, color="lt2" if dark else "dk2",
                  align="r", check=False)
        self.text(s, SW - MX - 1.1, 6.93, 1.1, 0.3, "PHILIPS", size=13, bold=True, color=fg, align="r",
                  spacing=150, check=False)
        if sd.get("footnotes"):
            fn = sd["footnotes"]
            fn = fn if isinstance(fn, list) else [fn]
            fy = sd.get("footnotes_y", 6.2)
            self.text(s, MX, fy, sd.get("footnotes_w", SW - 2 * MX), 6.85 - fy, fn, size=8,
                      color="lt2" if dark else "dk2", anchor="b", line=1.0)

    def notes(self, s, sd):
        parts = []
        if sd.get("notes"):
            parts.append(sd["notes"].strip())
        src = sd.get("src")
        if src:
            src = src if isinstance(src, list) else [src]
            parts.append("Sources : " + " ; ".join(self.expand_src(x) for x in src))
        if parts:
            s.notes_slide.notes_text_frame.text = "\n\n".join(parts)

    def expand_src(self, ref):
        m = re.fullmatch(r"([GJC])-(\d+)(?:-(\d+))?", str(ref))
        if not m:
            return str(ref)
        doc = {"G": "guide_du_marche_2025 (PDF)", "J": "presentation_JFR_2026 (PDF)",
               "C": "couverture 2026 (.pptx)"}[m.group(1)]
        a, b = int(m.group(2)), m.group(3)
        rng = f"{a}" if not b else f"{a} à {b}"
        if m.group(1) == "G":
            sl = f"{a + 1}" if not b else f"{a + 1} à {int(b) + 1}"
            return f"{doc}, page {rng} (n° de slide affiché {sl})"
        return f"{doc}, page {rng}"

    def collect_markers(self, sd):
        if sd.get("type") == "versus":
            names = [sd["philips"]["name"], sd["competitor"]["name"]] + \
                    ([sd["secondary"]["name"]] if sd.get("secondary") else [])
            for row in sd["rows"]:
                for name, v in zip(names, row["values"]):
                    if str(v) in (TODO, CHECK):
                        self.todos.append((self.slide_no, sd.get("title", ""), f"{v} {row['label']} – {name}"))
            sd = {k: v for k, v in sd.items() if k != "rows"}
        if sd.get("type") == "table":
            for row in sd["rows"]:
                for head, v in zip(sd["columns"], row):
                    if str(v) in (TODO, CHECK):
                        self.todos.append((self.slide_no, sd.get("title", ""),
                                           f"{v} {' → '.join(str(c) for c in row if str(c) not in (TODO, CHECK))}"))
            sd = {k: v for k, v in sd.items() if k != "rows"}

        def walk(v, path):
            if isinstance(v, dict):
                for k, x in v.items():
                    if k not in ("notes", "src", "todo_notes"):
                        walk(x, path + [k])
            elif isinstance(v, list):
                for x in v:
                    walk(x, path)
            elif isinstance(v, str):
                if path and path[-1] in ("rows",) and isinstance(v, str) and v in (TODO, CHECK):
                    self.todos.append((self.slide_no, sd.get("title", ""), v))
                    return
                for m in (TODO, CHECK):
                    if m in v:
                        ctx = v.replace(m, "").strip(" :–-") or "/".join(str(p) for p in path)
                        self.todos.append((self.slide_no, sd.get("title", ""), f"{m} {ctx}"))
        walk(sd, [])
        for t in sd.get("todo_notes", []):
            self.todos.append((self.slide_no, sd.get("title", ""), t))

    # ------------------------------------------------------------------ types de slides
    def s_cover(self, sd):
        s = self.prs.slides[0]
        self.slide_no = 1
        self.cur_title = "Couverture"
        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        for old, new in sd.get("replace", {}).items():
                            if old in r.text:
                                r.text = r.text.replace(old, new)
        self.notes(s, sd)
        self.ids["couverture"] = s

    def s_toc(self, sd):
        s = self.new_slide(sd)
        self.set_title(s, sd["title"], y=0.55)
        self.image(s, sd["image"], 7.9, 0, SW - 7.9, SH, focus=(0.55, 0.5))
        ov = self.rect(s, 7.9, 0, 2.2, SH, fill="dk1")
        self.gradient(ov, [(0, "dk1", 1), (1, "dk1", 0)], angle=0)
        secs = [x for x in self.content["sections"] if not x.get("front")]
        cols, tw, th, gx, gy = 2, 3.45, 1.12, 0.25, 0.22
        for i, sec in enumerate(secs):
            cx, cy = MX + (i % cols) * (tw + gx), 1.6 + (i // cols) * (th + gy)
            card = self.rect(s, cx, cy, tw, th, fill="dk2", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08,
                             name=f"Sommaire {sec['title']}")
            self.rect(s, cx + 0.25, cy + 0.28, 0.5, 0.5, fill=sec["color"], shape=MSO_SHAPE.OVAL)
            self.text(s, cx + 0.25, cy + 0.28, 0.5, 0.5, f"{i + 1}", size=16, bold=True,
                      color=self.P.on(sec["color"]), align="c", anchor="m", check=False)
            self.text(s, cx + 0.95, cy + 0.2, tw - 1.1, 0.42, sec["title"], size=15, bold=True)
            self.text(s, cx + 0.95, cy + 0.62, tw - 1.1, 0.4, sec.get("tagline", ""), size=10, color="lt2")
            self.link(card, sec["id"])
        self.chrome(s, sd, show_section=False)

    def s_opener(self, sd):
        """Ouverture de section immersive : image plein cadre + voile dégradé + titre XXL + chapitres."""
        s = self.new_slide(sd)
        sec = self.section
        self.image(s, sd["image"], 0, 0, SW, SH, focus=tuple(sd.get("focus", (0.7, 0.5))))
        ov = self.rect(s, 0, 0, SW, SH, fill="dk1")
        self.gradient(ov, [(0, "dk1", 1), (0.42, "dk1", 0.92), (0.75, "dk1", 0.15), (1, "dk1", 0)], angle=0)
        num = sd.get("number")
        if num:
            self.text(s, MX, 1.05, 3, 1.3, f"{num:02d}", size=80, bold=True, color=sec["color"], check=False)
        self.set_title(s, sd["title"], y=2.45, w=6.4, h=1.7, size=50)
        if sd.get("subtitle"):
            self.text(s, MX, 4.2, 5.6, 0.9, sd["subtitle"], size=16, color="lt2")
        if sd.get("badge_new"):
            self.badge_new(s, MX, 2.05)
        chips = sd.get("chapters", [])
        cx, cy = MX, 5.25
        for ch in chips:
            w = text_width(ch["label"], 11, True) + 0.45
            if cx + w > 7.3:
                cx, cy = MX, cy + 0.45
            c = self.rect(s, cx, cy, w, 0.34, line=sec["color"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5,
                          line_w=1)
            self.text(s, cx, cy, w, 0.34, ch["label"], size=11, bold=True, color="lt1", align="c", anchor="m",
                      check=False)
            self.link(c, ch["link"])
            cx += w + 0.15
        self.chrome(s, sd)

    def s_xxl(self, sd):
        """Chiffres clés géants (1 à 4) + image optionnelle."""
        s = self.new_slide(sd)
        self._allow_2l = True
        self.set_title(s, sd["title"], w=(7.0 if sd.get("image") else SW - 2 * MX), h=1.3)
        self._allow_2l = False
        shift = 0.5 * (self.title_lines - 1)
        stats = sd["stats"]
        has_img = bool(sd.get("image"))
        area_w = 7.0 if has_img else SW - 2 * MX
        if has_img:
            self.image(s, sd["image"], 7.9, 0, SW - 7.9, SH, mode=sd.get("image_mode", "cover"),
                       focus=tuple(sd.get("focus", (0.5, 0.5))))
            ov = self.rect(s, 7.9, 0, 1.6, SH, fill="dk1")
            self.gradient(ov, [(0, "dk1", 1), (1, "dk1", 0)])
        n = len(stats)
        per_row = sd.get("per_row") or (n if n <= 3 or (not has_img and n <= 4) else 2)
        rows = (n + per_row - 1) // per_row
        cw = area_w / per_row
        top = (2.0 if rows == 1 else 1.75) + shift + (0.3 if sd.get("lead") else 0)
        rh = 3.6 if rows == 1 else 2.1
        big = sd.get("size", 88 if (rows == 1 and per_row <= 2) else (66 if rows == 1 else 54))
        for i, st in enumerate(stats):
            x = MX + (i % per_row) * cw
            y = top + (i // per_row) * rh
            val = str(st["value"])
            unit = st.get("unit", "")
            vs = big
            while text_width(val, vs, True) + text_width(unit, vs * 0.4, True) + 0.1 > cw - 0.25 and vs > 30:
                vs -= 2
            self.text(s, x, y, cw - 0.2, vs / 72 * 1.15,
                      [[(val, {"size": vs, "bold": True, "color": "accent1"}),
                        (f" {unit}" if unit else "", {"size": int(vs * 0.4), "bold": True, "color": "accent1"})]],
                      check=False)
            ly = y + vs / 72 * 1.18
            self.text(s, x, ly, cw - 0.35, rh - vs / 72 * 1.2 - 0.15, [[(st["label"], {"bold": True})]]
                      + ([st["detail"]] if st.get("detail") else []), size=14, color="lt1", para_space=4)
        if sd.get("lead"):
            self.text(s, MX, 1.45 + shift, area_w, 0.5, sd["lead"], size=14, color="lt2")
        if has_img:
            sd.setdefault("footnotes_w", 7.0)
        self.chrome(s, sd)

    def s_spec(self, sd):
        """Fiche technique : visuel produit + grille de pictos (valeur + libellé)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"])
        if sd.get("badge_new"):
            self.badge_new(s, SW - MX - 0.95, 0.85)
        if sd.get("image"):
            if self.has_alpha(sd["image"]):
                halo = self.rect(s, 0.4, 1.9, 4.4, 4.4, fill="dk2", shape=MSO_SHAPE.OVAL)
                self.gradient(halo, [(0, "accent1/dk1/0.55", 1), (1, "dk1", 1)], angle=90)
                self.image(s, sd["image"], 0.45, 1.75, 4.3, 4.5, mode="contain")
            else:
                pic = self.image(s, sd["image"], MX, 1.85, 4.1, 4.1, focus=tuple(sd.get("focus", (0.5, 0.5))))
                self.round_pic(pic, 0.06)
        else:
            self.visual_placeholder(s, MX, 1.9, 4.0, 4.2, sd.get("image_hint", ""))
        items = sd["items"]
        cols = sd.get("cols", 3)
        x0, y0, cw = 5.2, 1.85, (SW - MX - 5.2) / cols
        rows = (len(items) + cols - 1) // cols
        rh = min(1.25, 4.45 / rows)
        for i, it in enumerate(items):
            x, y = x0 + (i % cols) * cw, y0 + (i // cols) * rh
            self.token(s, x, y + 0.05, 0.52, it.get("icon", "•"), size=10 if len(it.get("icon", "")) > 2 else 12)
            v = str(it["value"])
            label = it["label"]
            if CHECK in v:
                v = v.replace(CHECK, "").strip()
                label = f"{label}  {CHECK}"
            if v == TODO:
                self.todo_box(s, x + 0.65, y + 0.02, cw - 0.8, 0.36, size=10)
            else:
                self.text(s, x + 0.65, y, cw - 0.75, 0.45, v, size=it.get("size", 18), bold=True,
                          color="accent1")
            self.text(s, x + 0.65, y + 0.45, cw - 0.75, rh - 0.5, label, size=11, color="lt2")
        self.chrome(s, sd)

    def s_cards(self, sd):
        """Cartes produits groupées par famille (gamme)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"])
        if sd.get("lead"):
            self.text(s, MX, 1.45, SW - 2 * MX, 0.4, sd["lead"], size=14, color="lt2")
        groups = sd["groups"]
        total = sum(len(g["items"]) for g in groups)
        gap_g, gap_i = 0.3, 0.12
        avail = SW - 2 * MX - gap_g * (len(groups) - 1)
        unit = (avail - gap_i * (total - len(groups))) / total
        x, y, h = MX, 2.05, 4.0
        for g in groups:
            gw = unit * len(g["items"]) + gap_i * (len(g["items"]) - 1)
            col = g.get("color", "accent1")
            self.rect(s, x, y, gw, 0.42, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
            self.text(s, x, y, gw, 0.42, g["label"], size=12, bold=True, color=self.P.on(col), align="c",
                      anchor="m", check=False)
            if g.get("pillar"):
                self.text(s, x, y + 0.5, gw, 0.5, g["pillar"], size=10, color="lt2", align="c", italic=True)
            ix = x
            for it in g["items"]:
                cy = y + 1.05
                card = self.rect(s, ix, cy, unit, h - 1.05, fill="dk2", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                                 radius=0.06, name=f"Carte {it['name']}")
                if it.get("image") and self.has_alpha(it["image"]):
                    self.image(s, it["image"], ix + 0.08, cy + 0.15, unit - 0.16, 1.55, mode="contain")
                elif it.get("image"):
                    pic = self.image(s, it["image"], ix + 0.08, cy + 0.12, unit - 0.16, 1.6,
                                     focus=tuple(it.get("focus", (0.5, 0.55))))
                    self.round_pic(pic, 0.08)
                else:
                    self.visual_placeholder(s, ix + 0.1, cy + 0.15, unit - 0.2, 1.5, it.get("image_hint", ""))
                self.text(s, ix + 0.08, cy + 1.8, unit - 0.16, 0.65, it["name"], size=12, bold=True, align="c")
                if it.get("sub"):
                    self.text(s, ix + 0.08, cy + 2.4, unit - 0.16, 0.45, it["sub"], size=9, color="lt2",
                              align="c")
                if it.get("new"):
                    self.badge_new(s, ix + unit / 2 - 0.475, cy - 0.15)
                if it.get("link"):
                    self.link(card, it["link"])
                ix += unit + gap_i
            x += gw + gap_g
        self.chrome(s, sd)

    def s_fullbleed(self, sd):
        """Grand visuel plein cadre + message court."""
        s = self.new_slide(sd)
        self.image(s, sd["image"], 0, 0, SW, SH, focus=tuple(sd.get("focus", (0.5, 0.5))))
        side = sd.get("side", "l")
        ov = self.rect(s, 0, 0, SW, SH, fill="dk1")
        stops = [(0, "dk1", 0.95), (0.45, "dk1", 0.75), (0.7, "dk1", 0.0), (1, "dk1", 0)]
        self.gradient(ov, stops, angle=0 if side == "l" else 180)
        x = MX if side == "l" else SW - MX - 5.8
        self._allow_2l = True
        self.set_title(s, sd["title"], x=x, y=1.5, w=5.8, h=1.7, size=34)
        self._allow_2l = False
        if self.title_lines > 2:
            self.warnings.append(f"slide {self.slide_no}: titre plein cadre sur {self.title_lines} lignes")
        if sd.get("big"):
            self.text(s, x, 3.3, 5.4, 1.4, [[(sd["big"]["value"], {"size": 80, "bold": True, "color": "accent1"}),
                                             (" " + sd["big"].get("unit", ""), {"size": 32, "bold": True,
                                                                               "color": "accent1"})]],
                      check=False)
        if sd.get("body"):
            self.text(s, x, 4.75 if sd.get("big") else 3.4, 5.2, 1.4, sd["body"], size=15, color="lt1")
        self.chrome(s, sd)

    def s_overview(self, sd):
        """Vue d'ensemble : segments (lignes) × constructeurs (colonnes)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"], y=0.62, h=0.6, size=26)
        vendors, segs = sd["vendors"], sd["segments"]
        x0, y0 = MX + 1.45, 1.55
        cw = (SW - MX - x0) / len(vendors)
        hdr = 0.42
        for j, v in enumerate(vendors):
            col = "accent1" if v.get("philips") else "accent4"
            self.rect(s, x0 + j * cw + 0.04, y0, cw - 0.08, hdr, fill=col if v.get("philips") else "dk2",
                      line=None if v.get("philips") else "accent4", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
            self.text(s, x0 + j * cw, y0, cw, hdr, v["name"], size=12, bold=True,
                      color="dk1" if v.get("philips") else "lt1", align="c", anchor="m", check=False)
        avail = 6.75 - (y0 + hdr + 0.08)
        weights = [seg.get("weight", 1) for seg in segs]
        y = y0 + hdr + 0.08
        for seg, wt in zip(segs, weights):
            rh = avail * wt / sum(weights)
            band = self.rect(s, MX, y, SW - 2 * MX, rh - 0.06, fill="dk2", shape=MSO_SHAPE.RECTANGLE)
            self.gradient(band, [(0, "dk2", 1), (1, "dk1", 1)])
            self.text(s, MX + 0.12, y + 0.08, 1.25, 0.32, seg["name"], size=13, bold=True, color="accent1")
            self.text(s, MX + 0.12, y + 0.42, 1.25, rh - 0.5, seg.get("desc", ""), size=8, color="lt2", line=0.95)
            for j, v in enumerate(vendors):
                items = seg["cells"].get(v["key"], [])
                if not items:
                    continue
                n = len(items)
                ih = min(0.62, (rh - 0.14) / n - 0.05)
                for k, it in enumerate(items):
                    ix, iy = x0 + j * cw + 0.06, y + 0.04 + k * (ih + 0.05)
                    ph = v.get("philips")
                    fill = "accent1/dk2/0.15" if ph else "accent4/dk2/0.55"
                    if it.get("new2026"):
                        fill = "accent1"
                    box = self.rect(s, ix, iy, cw - 0.12, ih, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                                    radius=0.12, name=it["name"])
                    tcol = "dk1" if ph or it.get("new2026") else "lt1"
                    lines = [[(it["name"], {"bold": True, "size": 9.5})]]
                    if it.get("spec") and ih > 0.38:
                        lines.append([(it["spec"], {"size": 7.5})])
                    self.text(s, ix + 0.06, iy, cw - 0.24, ih, lines, color=tcol, anchor="m", line=0.9,
                              check=False)
                    if it.get("new2026"):
                        self.text(s, ix + cw - 0.75, iy + 0.02, 0.6, 0.2, "NOUVEAU", size=6.5, bold=True,
                                  color="dk1", align="r", check=False)
                    if it.get("link"):
                        self.link(box, it["link"])
            y += rh
        self.chrome(s, sd)

    def s_versus(self, sd):
        """Face-à-face technique : critères en lignes, Philips (cyan) vs concurrent (gris)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"], y=0.62, h=0.55, size=24)
        ph, cp = sd["philips"], sd["competitor"]
        sec = sd.get("secondary")
        lab_w = 2.6
        cols = [ph, cp] + ([sec] if sec else [])
        x_lab = MX
        cw = (SW - 2 * MX - lab_w - 1.9) / len(cols)
        x_cols = [x_lab + lab_w + i * cw for i in range(len(cols))]
        # visuels + noms
        img_w = 1.9
        self.image(s, ph.get("image"), SW - MX - img_w, 1.2, img_w, 1.3, mode="contain") if ph.get("image") \
            else None
        for i, c in enumerate(cols):
            philips = i == 0
            col = "accent1" if philips else ("accent4" if i == 1 else "accent4/dk1/0.4")
            self.rect(s, x_cols[i] + 0.05, 1.35, cw - 0.1, 0.62, fill=col if philips else "dk2",
                      line=None if philips else col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
            self.text(s, x_cols[i] + 0.05, 1.35, cw - 0.1, 0.62,
                      [[(c["name"], {"bold": True, "size": 13})],
                       [(c.get("vendor", ""), {"size": 9})]],
                      color="dk1" if philips else "lt1", align="c", anchor="m", line=0.95, check=False)
            if c.get("new"):
                self.badge_new(s, x_cols[i] + cw / 2 - 0.475, 1.12)
        self.text(s, x_lab, 1.5, lab_w, 0.4, "Critère", size=10, bold=True, color="lt2", anchor="m")
        rows = sd["rows"]
        y0, rh = 2.1, min(0.46, 4.0 / len(rows))
        for r_i, row in enumerate(rows):
            y = y0 + r_i * rh
            if r_i % 2 == 0:
                self.rect(s, MX, y, x_cols[-1] + cw - MX, rh, fill="dk2")
            self.token(s, x_lab + 0.05, y + (rh - 0.32) / 2, 0.32, row.get("icon", "•"), size=7)
            self.text(s, x_lab + 0.45, y, lab_w - 0.5, rh, row["label"], size=11, color="lt1", anchor="m",
                      check=False)
            vals = row["values"]
            adv = row.get("advantage") == "philips"
            for i, v in enumerate(vals):
                v = str(v)
                x = x_cols[i]
                if v == TODO:
                    self.todo_box(s, x + 0.25, y + 0.06, cw - 0.5, rh - 0.12, size=9)
                    continue
                col = "accent1" if i == 0 else "lt1"
                self.text(s, x + 0.1, y, cw - (0.55 if (i == 0 and adv) else 0.2), rh, v, size=11,
                          bold=(i == 0), color=col, align="c", anchor="m", check=False)
                if i == 0 and adv:
                    self.rect(s, x + cw - 0.42, y + (rh - 0.26) / 2, 0.26, 0.26, fill="accent1", shape=MSO_SHAPE.OVAL,
                              name="Avantage Philips")
                    self.text(s, x + cw - 0.42, y + (rh - 0.26) / 2, 0.26, 0.26, "✓", size=10, bold=True,
                              color="dk1", align="c", anchor="m", check=False)
        # légende
        ly = y0 + len(rows) * rh + 0.12
        self.rect(s, MX, ly + 0.04, 0.2, 0.2, fill="accent1", shape=MSO_SHAPE.OVAL)
        self.text(s, MX + 0.28, ly, 4.5, 0.3, "Avantage Philips (donnée sourcée)", size=9, color="lt2")
        self.rect(s, MX + 3.2, ly + 0.02, 0.5, 0.22, line="accent4", dash=True, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                  radius=0.3)
        self.text(s, MX + 3.8, ly, 4.5, 0.3, "Donnée absente des sources : à compléter", size=9, color="lt2")
        if sd.get("next"):
            n = self.rect(s, SW - MX - 2.6, ly - 0.02, 2.6, 0.34, line="accent1", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                          radius=0.5)
            self.text(s, SW - MX - 2.6, ly - 0.02, 2.6, 0.34, "Arguments & questions  ›", size=10, bold=True,
                      color="accent1", align="c", anchor="m", check=False)
            self.link(n, sd["next"])
        self.chrome(s, sd)

    def s_arguments(self, sd):
        """Arguments différenciants (3-4) + encadré « Questions à poser au client »."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"], y=0.62, h=0.55, size=24)
        if sd.get("lead"):
            self.text(s, MX, 1.2, 8.0, 0.4, sd["lead"], size=13, color="lt2")
        args = sd["arguments"]
        aw = 7.6
        ah = min(1.18, 4.75 / len(args) - 0.12)
        for i, a in enumerate(args):
            y = 1.75 + i * (ah + 0.12)
            self.rect(s, MX, y, aw, ah, fill="dk2", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
            self.token(s, MX + 0.2, y + (ah - 0.6) / 2, 0.6, a.get("icon", str(i + 1)), fill="accent1",
                       line="accent1", color="dk1", size=11)
            self.text(s, MX + 1.0, y + 0.12, aw - 1.2, 0.35, a["title"], size=15, bold=True, color="accent1")
            self.text(s, MX + 1.0, y + 0.47, aw - 1.2, ah - 0.55, a["text"], size=11, color="lt1", line=1.0)
        qx, qw = MX + aw + 0.3, SW - 2 * MX - aw - 0.3
        box = self.rect(s, qx, 1.75, qw, 4.75 - 0.12, fill="accent1/dk1/0.82", line="accent1", line_w=1.25,
                        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        self.text(s, qx + 0.25, 1.92, qw - 0.5, 0.4, "Questions à poser au client", size=15, bold=True,
                  color="accent1")
        qs = sd.get("questions") or [TODO]
        self.text(s, qx + 0.25, 2.45, qw - 0.5, 3.9, [[("?  ", {"bold": True, "color": "accent1"}), (q, {})]
                                                      for q in qs], size=12, color="lt1", para_space=8)
        if sd.get("back"):
            b = self.rect(s, SW - MX - 2.0, 1.2, 2.0, 0.34, line="accent1", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                          radius=0.5)
            self.text(s, SW - MX - 2.0, 1.2, 2.0, 0.34, "‹  Face-à-face", size=10, bold=True, color="accent1",
                      align="c", anchor="m", check=False)
            self.link(b, sd["back"])
        self.chrome(s, sd)

    def s_procons(self, sd):
        """Versus 2025 conservé : points forts du concurrent / points d'attention (données d'origine)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"], y=0.62, h=0.55, size=24)
        cols = [("Points forts du concurrent", sd.get("pros", []), "accent4"),
                ("Points d'attention", sd.get("cons", []), "accent1")]
        cw = (SW - 2 * MX - 0.3) / 2
        for i, (head, items, col) in enumerate(cols):
            x = MX + i * (cw + 0.3)
            self.rect(s, x, 1.45, cw, 5.0, fill="dk2", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
            self.token(s, x + 0.3, 1.65, 0.5, "+" if i == 0 else "!", fill=col, line=col, color="dk1", size=16)
            self.text(s, x + 0.95, 1.7, cw - 1.2, 0.45, head, size=16, bold=True, color=col)
            self.text(s, x + 0.35, 2.4, cw - 0.7, 3.9, [[("›  ", {"bold": True, "color": col}), (t, {})]
                                                         for t in items], size=13, para_space=6)
        if sd.get("back"):
            b = self.rect(s, SW - MX - 2.0, 0.85, 2.0, 0.34, line="accent1", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                          radius=0.5)
            self.text(s, SW - MX - 2.0, 0.85, 2.0, 0.34, "‹  Face-à-face", size=10, bold=True, color="accent1",
                      align="c", anchor="m", check=False)
            self.link(b, sd["back"])
        self.chrome(s, sd)

    def s_table(self, sd):
        """Tableau simple (ex. correspondance des noms)."""
        s = self.new_slide(sd)
        self.set_title(s, sd["title"])
        if sd.get("lead"):
            self.text(s, MX, 1.5, SW - 2 * MX, 0.45, sd["lead"], size=14, color="lt2")
        cols, rows = sd["columns"], sd["rows"]
        widths = sd.get("widths") or [1] * len(cols)
        tw = SW - 2 * MX
        ws = [tw * w / sum(widths) for w in widths]
        y, rh = 2.15, sd.get("row_h", 0.55)
        x = MX
        for c, w in zip(cols, ws):
            self.rect(s, x, y, w - 0.04, rh, fill="accent1")
            self.text(s, x + 0.15, y, w - 0.3, rh, c, size=13, bold=True, color="dk1", anchor="m", check=False)
            x += w
        for i, r in enumerate(rows):
            y += rh + 0.04
            x = MX
            for j, (c, w) in enumerate(zip(r, ws)):
                self.rect(s, x, y, w - 0.04, rh, fill="dk2" if i % 2 == 0 else "dk2/dk1/0.5")
                self.text(s, x + 0.15, y, w - 0.3, rh, str(c), size=13, bold=(j == 0),
                          color="lt1" if j else "accent1", anchor="m", check=False)
                x += w
        self.chrome(s, sd)

    # ------------------------------------------------------------------ assemblage
    def build(self):
        for sec in self.content["sections"]:
            sec.setdefault("short", sec["title"])
        for i, sec in enumerate(self.content["sections"]):
            self.section = None if sec.get("front") else sec
            for j, sd in enumerate(sec["slides"]):
                if j == 0 and not sec.get("front"):
                    sd.setdefault("id", sec["id"])
                getattr(self, "s_" + sd["type"])(sd)
        for shape, target in self.links:
            if target not in self.ids:
                self.warnings.append(f"lien vers « {target} » sans cible")
                continue
            shape.click_action.target_slide = self.ids[target]
        self.prs.core_properties.title = self.meta["title"]
        self.prs.save(self.out_path)
        self.write_todo()
        print(f"{len(self.prs.slides)} slides → {self.out_path}")
        for w in self.warnings:
            print("ATTENTION", w)
        return not self.warnings

    def write_todo(self):
        path = self.out_path.parent / "A_COMPLETER.md"
        lines = ["# À compléter – Guide du marché CT 2026", "",
                 "Généré automatiquement par `tools/build.py` à partir de `content.yaml` "
                 "(marqueurs [À COMPLÉTER], [À VÉRIFIER], [VISUEL À AJOUTER]).", ""]
        lines += self.content.get("todo_preamble", [])
        lines += ["", "| Slide | Titre | Élément |", "|---|---|---|"]
        seen = set()
        for n, t, item in self.todos:
            key = (n, item)
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"| {n} | {t} | {item} |")
        path.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", default=str(ROOT / "livrables" / "content.yaml"))
    ap.add_argument("--out", default=str(ROOT / "livrables" / "Guide_du_marche_CT_2026.pptx"))
    a = ap.parse_args()
    ok = Builder(a.content, a.out).build()
    sys.exit(0 if ok else 1)
