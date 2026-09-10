#!/usr/bin/env python3
"""Swap pptxgenjs's grey play-button placeholder for a real poster frame.

pptxgenjs has no poster option for addMedia, so every embedded video ships with
a generic grey placeholder. This locates the poster image belonging to each
embedded video (via the <p:pic> that carries the videoFile reference) and
replaces that part with media/poster.png.

Usage: python3 postbuild.py deck.pptx media/poster.png
"""
import re
import shutil
import sys
import zipfile

import defusedxml.minidom  # noqa: F401  (kept: parsing OOXML elsewhere needs the safe parser)


def find_poster_parts(z):
    """Return the set of zip part names used as poster images for videos."""
    posters = set()
    for name in z.namelist():
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
        if not m:
            continue
        slide = z.read(name).decode("utf-8")
        rels_name = f"ppt/slides/_rels/slide{m.group(1)}.xml.rels"
        try:
            rels = z.read(rels_name).decode("utf-8")
        except KeyError:
            continue
        rel_target = dict(
            re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels)
        )
        # each <p:pic> holding <a:videoFile> also holds the poster in <a:blip r:embed>
        for pic in re.findall(r"<p:pic>.*?</p:pic>", slide, re.S):
            if "videoFile" not in pic:
                continue
            embed = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', pic)
            if not embed:
                continue
            target = rel_target.get(embed.group(1), "")
            if target.startswith("../"):
                target = "ppt/" + target[3:]
            posters.add(target)
    return posters


def main(deck, poster):
    with zipfile.ZipFile(deck) as z:
        parts = find_poster_parts(z)
        if not parts:
            print("no embedded video poster found; nothing to do")
            return
        items = z.infolist()
        payload = {i.filename: z.read(i.filename) for i in items}

    with open(poster, "rb") as fh:
        poster_bytes = fh.read()
    for p in parts:
        payload[p] = poster_bytes

    tmp = deck + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for i in items:
            out.writestr(i, payload[i.filename])
    shutil.move(tmp, deck)
    print(f"replaced {len(parts)} video poster(s): {', '.join(sorted(parts))}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
