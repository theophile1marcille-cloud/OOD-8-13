#!/usr/bin/env python3
"""Allège un .pptx : les images PNG opaques passent en JPEG, et toutes les images sont limitées à MAX_SIDE px.

Usage : python3 tools/compact_pptx.py entree.pptx [sortie.pptx]
Le contenu des slides n'est pas modifié ; seules les images (ppt/media) et leurs références sont réécrites.
"""
import io
import re
import sys
import zipfile

from PIL import Image

MAX_SIDE = 2000
JPEG_QUALITY = 85


def opaque(im):
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        alpha = im.convert("RGBA").getchannel("A")
        return alpha.getextrema()[0] >= 250
    return True


def compact(src, dst):
    zin = zipfile.ZipFile(src)
    renames, blobs = {}, {}
    for info in zin.infolist():
        name = info.filename
        data = zin.read(name)
        if name.startswith("ppt/media/") and name.lower().endswith(".png"):
            im = Image.open(io.BytesIO(data))
            im.load()
            if max(im.size) > MAX_SIDE:
                im.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
            out = io.BytesIO()
            if opaque(im):
                im.convert("RGB").save(out, "JPEG", quality=JPEG_QUALITY, optimize=True)
                new = name[:-4] + ".jpeg"
                if out.tell() < len(data):
                    renames[name.split("/")[-1]] = new.split("/")[-1]
                    blobs[new] = out.getvalue()
                    continue
            else:
                im.save(out, "PNG", optimize=True)
                if out.tell() < len(data):
                    data = out.getvalue()
        blobs[name] = data
    keys = sorted(renames, key=len, reverse=True)
    pat = re.compile("(?<=media/)(" + "|".join(re.escape(k) for k in keys) + ")") if renames else None
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for name, data in blobs.items():
        if pat and (name.endswith(".rels") or name.endswith(".xml")):
            text = data.decode("utf-8")
            text = pat.sub(lambda m: renames[m.group(1)], text)
            if name == "[Content_Types].xml" and 'Extension="jpeg"' not in text:
                text = text.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
            data = text.encode("utf-8")
        zout.writestr(name, data)
    zout.close()
    return len(renames)


if __name__ == "__main__":
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    tmp = dst + ".tmp"
    n = compact(src, tmp)
    import os
    os.replace(tmp, dst)
    print(f"{n} images converties en JPEG → {dst} ({os.path.getsize(dst) / 1e6:.1f} Mo)")
