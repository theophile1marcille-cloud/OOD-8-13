"""Étape 1 – Analyse des sources (rejouable).

Les deux sources ont été fournies en PDF (export des .pptx), découpées en parties.
Pour chaque document : texte par page, rendu PNG, images embarquées
(page d'origine, dimensions natives, résolution effective), dédoublonnées.
Le thème est lu dans le seul .pptx disponible (couverture 2026).
Usage : python3 tools/analyse.py
"""
import hashlib, json, re, subprocess, zipfile
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
SRC, OUT = ROOT / "sources", ROOT / "analyse"
DOCS = {"guide_2025": SRC / "guide_2025_pdf", "jfr_2026": SRC / "jfr_2026_pdf"}
MIN_SIDE = 120  # ignore pictos/masques plus petits

def parts(d):
    return sorted(d.glob("p*.pdf"))

def npages(pdf):
    out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
    return int(re.search(r"Pages:\s+(\d+)", out).group(1))

def extract_doc(name, folder):
    img_dir = OUT / "images" / name
    img_dir.mkdir(parents=True, exist_ok=True)
    pages, images, seen, offset = [], [], {}, 0
    for pdf in parts(folder):
        n = npages(pdf)
        listing = subprocess.run(["pdfimages", "-list", str(pdf)], capture_output=True, text=True).stdout.splitlines()[2:]
        meta = {}
        for line in listing:
            c = line.split()
            if len(c) > 13:
                meta[int(c[1])] = {"type": c[2], "ppi_x": c[12]}
        tmp = img_dir / "_tmp"
        tmp.mkdir(exist_ok=True)
        subprocess.run(["pdfimages", "-png", "-p", str(pdf), str(tmp / "i")], check=True)
        for f in sorted(tmp.glob("i-*.png")):
            _, pg, num = f.stem.split("-")
            page, num = offset + int(pg), int(num)
            if meta.get(num, {}).get("type") == "smask":
                f.unlink(); continue
            with Image.open(f) as im:
                w, h = im.size
            if min(w, h) < MIN_SIDE:
                f.unlink(); continue
            digest = hashlib.md5(f.read_bytes()).hexdigest()[:12]
            if digest in seen:
                seen[digest]["pages"].append(page); f.unlink(); continue
            dest = img_dir / f"p{page:03d}_{digest}.png"
            f.rename(dest)
            rec = {"file": str(dest.relative_to(OUT)), "pages": [page], "w": w, "h": h,
                   "ppi": meta.get(num, {}).get("ppi_x")}
            seen[digest] = rec; images.append(rec)
        tmp.rmdir()
        for i in range(1, n + 1):
            txt = subprocess.run(["pdftotext", "-layout", "-f", str(i), "-l", str(i), str(pdf), "-"],
                                 capture_output=True, text=True).stdout
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            pages.append({"page": offset + i, "part": pdf.name, "text": "\n".join(lines),
                          "copyright": [l for l in lines if "Koninklijke" in l]})
        offset += n
    return pages, images

def theme_from_pptx(pptx):
    z = zipfile.ZipFile(pptx)
    xml = z.read("ppt/theme/theme1.xml").decode()
    colors = dict(re.findall(r"<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)><a:srgbClr val=\"(\w{6})\"", xml))
    fonts = {"major": re.search(r"<a:majorFont><a:latin typeface=\"([^\"]*)\"", xml).group(1),
             "minor": re.search(r"<a:minorFont><a:latin typeface=\"([^\"]*)\"", xml).group(1)}
    layouts = []
    for n in sorted(x for x in z.namelist() if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", x)):
        lx = z.read(n).decode()
        layouts.append({"part": n, "name": (re.search(r"<p:cSld name=\"([^\"]*)\"", lx) or [None, ""])[1],
                        "placeholders": re.findall(r"<p:ph(?: type=\"(\w+)\")?", lx)})
    sz = re.search(r"<p:sldSz cx=\"(\d+)\" cy=\"(\d+)\"", z.read("ppt/presentation.xml").decode())
    return {"source": pptx.name, "theme_name": re.search(r"<a:theme [^>]*name=\"([^\"]*)\"", xml).group(1),
            "colors": colors, "fonts": fonts, "layouts": layouts,
            "slide_size_in": [int(sz.group(1)) / 914400, int(sz.group(2)) / 914400]}

def main():
    summary = {}
    for name, folder in DOCS.items():
        pages, images = extract_doc(name, folder)
        (OUT / f"{name}_pages.json").write_text(json.dumps(pages, ensure_ascii=False, indent=1))
        (OUT / f"{name}_images.json").write_text(json.dumps(images, ensure_ascii=False, indent=1))
        with open(OUT / f"{name}_texte.md", "w") as md:
            md.write(f"# {name} – texte page par page\n\n")
            for p in pages:
                md.write(f"## Page {p['page']}\n\n```\n{p['text']}\n```\n\n")
        summary[name] = {"pages": len(pages), "images_uniques": len(images)}
    theme = theme_from_pptx(SRC / "guide_du_marche_2025.pptx")
    (OUT / "theme.json").write_text(json.dumps(theme, ensure_ascii=False, indent=1))
    summary["theme"] = theme
    (OUT / "resume.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1))
    print(json.dumps(summary, ensure_ascii=False, indent=1))

if __name__ == "__main__":
    main()
