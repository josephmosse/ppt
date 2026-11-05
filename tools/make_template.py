"""Create templates/irrbb_template.pptx: a 16:9 starter template you can replace with your own."""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "templates" / "irrbb_template.pptx"
KEEP = {"Title Slide", "Title and Content", "Two Content", "Title Only"}
THEME_COLOURS = {
    "dk1": "1E2A32", "lt1": "FFFFFF", "dk2": "12343B", "lt2": "E9EEF0",
    "accent1": "0F6E78", "accent2": "E07B00", "accent3": "4B3F8F", "accent4": "8E9AA6",
    "accent5": "5F6F7A", "accent6": "2F7ED8", "hlink": "0F6E78", "folHlink": "4B3F8F",
}
FONT = "Calibri"
MARGIN = Inches(0.6)
WIDTH, HEIGHT = Inches(13.333), Inches(7.5)
CONTENT_WIDTH = WIDTH - 2 * MARGIN
TITLE_TOP, TITLE_HEIGHT = Inches(0.4), Inches(1.05)
BODY_TOP, BODY_BOTTOM = Inches(1.6), Inches(6.75)
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def set_theme(prs):
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    for key, value in THEME_COLOURS.items():
        slot = theme.find(f".//{{{DRAWING_NS}}}clrScheme/{{{DRAWING_NS}}}{key}")
        for child in list(slot):
            slot.remove(child)
        etree.SubElement(slot, f"{{{DRAWING_NS}}}srgbClr", val=value)
    theme.find(f".//{{{DRAWING_NS}}}clrScheme").set("name", "IRRBB")
    fonts = theme.find(f".//{{{DRAWING_NS}}}fontScheme")
    fonts.set("name", "IRRBB")
    for group in ("majorFont", "minorFont"):
        fonts.find(f"{{{DRAWING_NS}}}{group}/{{{DRAWING_NS}}}latin").set("typeface", FONT)
    part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)


def place(shape, left, top, width, height):
    shape.left, shape.top, shape.width, shape.height = int(left), int(top), int(width), int(height)


def by_type(container, *types):
    return [p for p in container.placeholders if p.element.ph.get("type", "obj") in types]


def set_master_styles(master):
    styles = master._element.find(qn("p:txStyles"))
    title = styles.find(qn("p:titleStyle")).find(qn("a:lvl1pPr"))
    title.set("algn", "l")
    default = title.find(qn("a:defRPr"))
    default.set("sz", "2600")
    default.set("b", "1")
    fill = default.find(qn("a:solidFill"))
    for child in list(fill):
        fill.remove(child)
    etree.SubElement(fill, qn("a:schemeClr"), val="tx2")
    sizes = {"a:lvl1pPr": "1600", "a:lvl2pPr": "1400", "a:lvl3pPr": "1400", "a:lvl4pPr": "1200", "a:lvl5pPr": "1200"}
    body = styles.find(qn("p:bodyStyle"))
    for tag, size in sizes.items():
        body.find(qn(tag)).find(qn("a:defRPr")).set("sz", size)

    place(by_type(master, "title")[0], MARGIN, TITLE_TOP, CONTENT_WIDTH, TITLE_HEIGHT)
    place(by_type(master, "body")[0], MARGIN, BODY_TOP, CONTENT_WIDTH, BODY_BOTTOM - BODY_TOP)
    by_type(master, "title")[0].element.txBody.find(qn("a:bodyPr")).set("anchor", "t")
    place(by_type(master, "dt")[0], MARGIN, Inches(6.95), Inches(2), Inches(0.35))
    place(by_type(master, "ftr")[0], MARGIN, Inches(6.95), Inches(9), Inches(0.35))
    place(by_type(master, "sldNum")[0], WIDTH - MARGIN - Inches(1.5), Inches(6.95), Inches(1.5), Inches(0.35))
    # No inset, so titles, footers and text line up with the edges of tables and cards.
    for placeholder in by_type(master, "title", "body", "ftr"):
        placeholder.element.txBody.find(qn("a:bodyPr")).set("lIns", "0")
    by_type(master, "sldNum")[0].element.txBody.find(qn("a:bodyPr")).set("rIns", "0")
    for placeholder in by_type(master, "ftr"):
        level = placeholder.element.txBody.find(qn("a:lstStyle")).find(qn("a:lvl1pPr"))
        level.set("algn", "l")


def inherit_text_style(placeholder):
    """Drop the layout's size overrides so body text follows the master."""
    style = placeholder.element.txBody.find(qn("a:lstStyle"))
    for child in list(style):
        style.remove(child)


def inherit_geometry(placeholder):
    """Drop the layout's own position so the placeholder follows the master."""
    transform = placeholder.element.spPr.find(qn("a:xfrm"))
    if transform is not None:
        placeholder.element.spPr.remove(transform)


def set_text_style(placeholder, size, bold, colour, align="l"):
    style = placeholder.element.txBody.find(qn("a:lstStyle"))
    for child in list(style):
        style.remove(child)
    level = etree.SubElement(style, qn("a:lvl1pPr"), marL="0", indent="0", algn=align)
    etree.SubElement(level, qn("a:buNone"))
    run = etree.SubElement(level, qn("a:defRPr"), sz=str(size), b="1" if bold else "0")
    fill = etree.SubElement(run, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr"), val=colour)


def set_layouts(prs):
    for layout in list(prs.slide_layouts):
        if layout.name not in KEEP:
            prs.slide_layouts.remove(layout)
    for layout in prs.slide_layouts:
        for placeholder in by_type(layout, "dt", "ftr", "sldNum"):
            inherit_geometry(placeholder)
        if layout.name != "Title Slide":
            for placeholder in by_type(layout, "title"):
                inherit_geometry(placeholder)
        if layout.name == "Title and Content":
            inherit_geometry(by_type(layout, "obj")[0])
        elif layout.name == "Two Content":
            left, right = sorted(by_type(layout, "obj"), key=lambda p: p.placeholder_format.idx)
            height = BODY_BOTTOM - BODY_TOP
            chart_width = Inches(7.5)
            place(left, MARGIN, BODY_TOP, chart_width, height)
            place(right, MARGIN + chart_width + Inches(0.35), BODY_TOP, CONTENT_WIDTH - chart_width - Inches(0.35), height)
            inherit_text_style(left)
            inherit_text_style(right)
        elif layout.name == "Title Slide":
            layout.background.fill.solid()
            layout.background.fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_2
            title, subtitle = by_type(layout, "ctrTitle")[0], by_type(layout, "subTitle")[0]
            place(title, Inches(0.9), Inches(2.2), Inches(11.5), Inches(1.7))
            place(subtitle, Inches(0.9), Inches(4.05), Inches(11.5), Inches(1.2))
            title.element.txBody.find(qn("a:bodyPr")).set("anchor", "b")
            set_text_style(title, 4400, True, "bg1")
            set_text_style(subtitle, 1800, False, "bg2")
            for placeholder in by_type(layout, "dt", "ftr", "sldNum"):
                placeholder.element.getparent().remove(placeholder.element)


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = WIDTH, HEIGHT
    set_theme(prs)
    set_master_styles(prs.slide_master)
    set_layouts(prs)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
