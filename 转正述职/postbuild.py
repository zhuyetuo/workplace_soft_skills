#!/usr/bin/env python3
"""Post-process the generated deck for things pptxgenjs cannot express.

1. Swap the grey play-button placeholder for a real poster frame.
2. Mark the blow-up slides hidden, so they are reachable only by the
   thumbnail hyperlinks and never appear in the normal run.
3. Set embedded video to play full screen (needs a <p:timing> tree, which
   pptxgenjs does not emit).

Usage: python3 postbuild.py deck.pptx media/poster.png [hidden_slide_numbers]
       e.g. python3 postbuild.py deck.pptx media/poster.png 17,18
"""
import re
import shutil
import sys
import zipfile

TIMING = (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
    'nodeType="tmRoot"><p:childTnLst><p:video><p:cMediaNode vol="80000" '
    'fullScrn="1"><p:cTn id="2" fill="hold" display="0"><p:stCondLst>'
    '<p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl>'
    '<p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:video>'
    '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
)


def dedupe_shape_ids(payload):
    """pptxgenjs can emit the same <p:cNvPr id> twice on one slide (addMedia
    reuses the counter after addImage). PowerPoint refuses to open such a file,
    so give every duplicate a fresh id. Runs before the timing injection, which
    targets the video by id."""
    fixed = []
    for name in list(payload):
        if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
            continue
        x = payload[name].decode("utf-8")
        ids = [int(i) for i in re.findall(r'<p:cNvPr[^>]*\bid="(\d+)"', x)]
        if len(ids) == len(set(ids)):
            continue
        seen, nxt, out, pos = set(), max(ids) + 1, [], 0
        for m in re.finditer(r'(<p:cNvPr[^>]*\bid=")(\d+)(")', x):
            cur = int(m.group(2))
            out.append(x[pos:m.start()])
            if cur in seen:
                out.append(m.group(1) + str(nxt) + m.group(3))
                nxt += 1
            else:
                seen.add(cur)
                out.append(m.group(0))
            pos = m.end()
        out.append(x[pos:])
        payload[name] = "".join(out).encode("utf-8")
        fixed.append(name.rsplit("/", 1)[-1])
    return fixed


def poster_parts(payload):
    """Zip part names used as poster images for embedded videos."""
    out = set()
    for name, data in payload.items():
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
        if not m:
            continue
        slide = data.decode("utf-8")
        try:
            rels = payload[f"ppt/slides/_rels/slide{m.group(1)}.xml.rels"].decode("utf-8")
        except KeyError:
            continue
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        for pic in re.findall(r"<p:pic>.*?</p:pic>", slide, re.S):
            if "videoFile" not in pic:
                continue
            embed = re.search(r'<a:blip[^>]*r:embed="([^"]+)"', pic)
            if not embed:
                continue
            t = targets.get(embed.group(1), "")
            out.add("ppt/" + t[3:] if t.startswith("../") else t)
    return out


def set_video_fullscreen(payload):
    """Append a timing tree marking each embedded video as full-screen playback."""
    done = []
    for name in list(payload):
        if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
            continue
        slide = payload[name].decode("utf-8")
        if "videoFile" not in slide or "<p:timing>" in slide:
            continue
        pic = re.search(r"<p:pic>(?:(?!</p:pic>).)*videoFile(?:(?!</p:pic>).)*</p:pic>", slide, re.S)
        if not pic:
            continue
        spid = re.search(r'<p:cNvPr[^>]*id="(\d+)"', pic.group(0))
        if not spid:
            continue
        slide = slide.replace("</p:sld>", TIMING.format(spid=spid.group(1)) + "</p:sld>")
        payload[name] = slide.encode("utf-8")
        done.append(name)
    return done


def hide_slides(payload, numbers):
    done = []
    for n in numbers:
        key = f"ppt/slides/slide{n}.xml"
        if key not in payload:
            continue
        s = payload[key].decode("utf-8")
        if 'show="0"' in s:
            continue
        s = re.sub(r"(<p:sld\b[^>]*?)(\s*>)", r'\1 show="0"\2', s, count=1)
        payload[key] = s.encode("utf-8")
        done.append(n)
    return done


def main(deck, poster, hidden):
    with zipfile.ZipFile(deck) as z:
        items = z.infolist()
        payload = {i.filename: z.read(i.filename) for i in items}

    deduped = dedupe_shape_ids(payload)

    with open(poster, "rb") as fh:
        poster_bytes = fh.read()
    posters = poster_parts(payload)
    for p in posters:
        payload[p] = poster_bytes

    vids = set_video_fullscreen(payload)
    hid = hide_slides(payload, hidden)

    tmp = deck + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for i in items:
            out.writestr(i, payload[i.filename])
    shutil.move(tmp, deck)
    print(f"deduped shape ids: {deduped or 'none'} | poster: {len(posters)} "
          f"| fullscreen video: {len(vids)} | hidden slides: {hid}")


if __name__ == "__main__":
    nums = []
    if len(sys.argv) > 3:
        nums = [int(x) for x in sys.argv[3].split(",") if x.strip()]
    main(sys.argv[1], sys.argv[2], nums)
