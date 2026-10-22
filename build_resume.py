#!/usr/bin/env python3
"""Build Zach Yaninek's resume on the MIT Sloan MBAn 2027 template.

Reads resume content from a JSON file and writes a .docx that reuses the
template's styles, bullet numbering, fonts and margins unchanged; only the
document body is regenerated, following the template's own paragraph
patterns. Optionally converts the result to PDF with LibreOffice.

The phone number is never stored in this repository. Pass it at build time:

    python3 build_resume.py --phone 555-555-5555 --out private/resume.docx --pdf
"""

import argparse
import datetime
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = HERE / "template" / "MBAn-Resume_Template_2027.docx"
DEFAULT_CONTENT = HERE / "resume.json"

# Template geometry (twips). Letter page, 0.65" side margins.
TEXT_WIDTH = 12240 - 2 * 936
HEADING_TAB = 10250  # where the template's underlined section rules end
RIGHT_TAB = 10250    # right edge for locations and dates, flush with the rules
ENTRY_INDENT = 140   # template indent for org, degree and title lines
SZ = 20              # 10pt body text, as in the template
NAME_SZ = 22         # 11pt bold name, as in the template


def run(text, bold=False, italic=False, underline=False, sz=SZ, tab_before=False):
    props = ""
    if bold:
        props += "<w:b/>"
    if italic:
        props += "<w:i/>"
    if sz:
        props += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if underline:
        props += '<w:u w:val="single"/>'
    body = "<w:tab/>" if tab_before else ""
    if text:
        body += f'<w:t xml:space="preserve">{escape(text)}</w:t>'
    return f"<w:r><w:rPr>{props}</w:rPr>{body}</w:r>"


def styled_runs(text, bold=False, italic=False):
    """Split *italic* spans (used for species names) into separate runs."""
    out = []
    for i, part in enumerate(re.split(r"\*([^*]+)\*", text)):
        if part:
            out.append(run(part, bold=bold, italic=italic or i % 2 == 1))
    return "".join(out)


def para(runs, ppr=""):
    return f"<w:p><w:pPr>{ppr}<w:rPr><w:sz w:val=\"{SZ}\"/></w:rPr></w:pPr>{runs}</w:p>"


def tabbed(left_runs, right_text, gap=0, keep_next=True):
    """A line with text at the left and a right-aligned location or date."""
    ppr = ("<w:keepNext/>" if keep_next else "")
    ppr += f'<w:tabs><w:tab w:val="right" w:pos="{RIGHT_TAB}"/></w:tabs>'
    ppr += f'<w:spacing w:before="{gap}"/>' if gap else ""
    ppr += f'<w:ind w:left="{ENTRY_INDENT}"/>'
    right = run(right_text, tab_before=True) if right_text else ""
    return para(left_runs + right, ppr)


def centered(runs, gap):
    ppr = ('<w:pStyle w:val="BodyText"/>'
           f'<w:spacing w:before="{gap}"/>'
           '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/>')
    return para(runs, ppr)


def heading(title, gap):
    ppr = ('<w:pStyle w:val="Heading2"/><w:keepNext/>'
           f'<w:tabs><w:tab w:val="left" w:pos="{HEADING_TAB}"/></w:tabs>'
           f'<w:spacing w:before="{gap}" w:after="20"/>')
    runs = run(title, bold=True, underline=True) + run("", bold=True, underline=True, tab_before=True)
    return f"<w:p><w:pPr>{ppr}<w:rPr><w:u w:val=\"single\"/></w:rPr></w:pPr>{runs}</w:p>"


def bullet(text):
    ppr = ('<w:pStyle w:val="ListParagraph"/>'
           '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
           '<w:tabs><w:tab w:val="left" w:pos="320"/></w:tabs>'
           f'<w:ind w:left="320" w:right="{TEXT_WIDTH - RIGHT_TAB}" w:hanging="180"/>')
    return para(styled_runs(text), ppr)


def title_runs(title):
    """Bold italic title; a trailing parenthetical stays italic but not bold."""
    main, sep, rest = title.partition(" (")
    out = run(main, bold=True, italic=True)
    if sep:
        out += run(" (" + rest, italic=True)
    return out


def build_body(content, phone):
    layout = content.get("layout", {})
    section_gap = layout.get("section_gap", 200)
    entry_gap = layout.get("entry_gap", 120)
    h = content["header"]

    contact = h["contact_line"]
    if phone:
        contact = f"{phone} | {contact}"

    xml = [
        centered(run(h["name"], bold=True, sz=NAME_SZ), 4),
        centered(run(h["location"]), 10),
        centered(run(contact), 10),
        heading("EDUCATION", section_gap),
    ]
    for i, e in enumerate(content["education"]):
        xml.append(tabbed(run(e["org"], bold=True), e["location"], gap=entry_gap if i else 40))
        xml.append(tabbed(styled_runs(e["degree_line"], italic=True), e["dates"]))
        xml += [bullet(b) for b in e["bullets"]]

    xml.append(heading("TECHNICAL SKILLS", section_gap))
    xml += [bullet(b) for b in content["technical_skills"]]

    xml.append(heading("EXPERIENCE", section_gap))
    for i, e in enumerate(content["experience"]):
        xml.append(tabbed(run(e["org"], bold=True), e["location"], gap=entry_gap if i else 40))
        if e.get("description"):
            xml.append(tabbed(styled_runs(e["description"]), e["dates"]))
            xml.append(tabbed(title_runs(e["title"]), ""))
        else:
            xml.append(tabbed(title_runs(e["title"]), e["dates"]))
        xml += [bullet(b) for b in e["bullets"]]

    xml.append(heading("ADDITIONAL INFORMATION", section_gap))
    xml += [bullet(b) for b in content["additional"]]
    return "".join(xml)


def core_props(name):
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{escape(name)} Resume</dc:title><dc:creator>{escape(name)}</dc:creator>"
        f"<cp:lastModifiedBy>{escape(name)}</cp:lastModifiedBy><cp:revision>1</cp:revision>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def build_docx(template, content, phone, out):
    with zipfile.ZipFile(template) as zin:
        doc = zin.read("word/document.xml").decode("utf-8")
        # Keep the template's document element (namespaces) and section
        # properties (page size, margins); replace everything in between.
        start = doc.index("<w:body>") + len("<w:body>")
        sect = doc.index("<w:sectPr", start)
        new_doc = doc[:start] + build_body(content, phone) + doc[sect:]
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    data = new_doc.encode("utf-8")
                elif item.filename == "docProps/core.xml":
                    data = core_props(content["header"]["name"].title()).encode("utf-8")
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)


def to_pdf(docx):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("LibreOffice (soffice) not found; open the .docx in Word and Save as PDF instead")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx.parent), str(docx)],
                   check=True, stdout=subprocess.DEVNULL)
    return docx.with_suffix(".pdf")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--content", type=Path, default=DEFAULT_CONTENT)
    ap.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    ap.add_argument("--out", type=Path, default=HERE / "Yaninek_Zachary_MBAn2027_Resume.docx")
    ap.add_argument("--phone", help="phone number for the header; omitted from the public build")
    ap.add_argument("--pdf", action="store_true", help="also convert to PDF with LibreOffice")
    args = ap.parse_args()

    content = json.loads(args.content.read_text())
    build_docx(args.template, content, args.phone, args.out)
    print(args.out)
    if args.pdf:
        print(to_pdf(args.out))


if __name__ == "__main__":
    main()
