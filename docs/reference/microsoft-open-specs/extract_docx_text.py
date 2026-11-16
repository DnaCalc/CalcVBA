"""Extract the text of a Microsoft Open Specifications .docx.

Usage: python -I extract_docx_text.py <in.docx> <out.txt> [<out.images.tsv>]

Writes one line per non-empty paragraph of word/document.xml (runs joined; tabs and line breaks
dropped; table cells are separate paragraphs; LF line ends). This reproduces the earlier
extractions in Foundation/.tmp/spec_scan line for line, so line numbers cited in CalcVBA
documents (for example "MS-VBAL ll. 2793-2794") stay valid.

Several rules in the specs (formulas, some tables) are drawings whose meaning is only in the
picture's alt text. With a third argument, those are listed in a TSV: the .txt line the image
sits on or follows, the picture's title, and its description.
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def clean(text):
    return " ".join(text.split())


def main(src, dst, images_dst=None):
    with zipfile.ZipFile(src) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    lines, images = [], []
    for p in root.iter(W + "p"):
        text = "".join(t.text or "" for t in p.iter(W + "t"))
        if text:
            lines.append(text)
        for pr in p.iter(WP + "docPr"):
            title, descr = pr.get("title", ""), pr.get("descr", "")
            if title or descr:
                images.append((len(lines), clean(title), clean(descr)))
    with open(dst, "w", encoding="utf-8", newline="\n") as out:
        out.writelines(line + "\n" for line in lines)
    if images_dst:
        with open(images_dst, "w", encoding="utf-8", newline="\n") as out:
            out.write("line\ttitle\tdescription\n")
            out.writelines(f"{n}\t{t}\t{d}\n" for n, t, d in images)


if __name__ == "__main__":
    main(*sys.argv[1:4])
