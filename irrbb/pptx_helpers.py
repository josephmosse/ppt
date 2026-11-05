"""Template-agnostic python-pptx helpers.

Geometry comes from the template's placeholders and colours from its theme, so a
different template restyles the deck without code changes. Only status colours are
fixed, because red and green carry meaning whatever the brand.
"""
import copy
import math

from lxml import etree
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_LINE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .calc import EARLY_WARNING, LIMIT_BREACH, SOT_BREACH, WITHIN

STATUS_COLOURS = {
    WITHIN: RGBColor(0x2E, 0x7D, 0x32),
    EARLY_WARNING: RGBColor(0xB2, 0x6A, 0x00),
    LIMIT_BREACH: RGBColor(0xC6, 0x28, 0x28),
    SOT_BREACH: RGBColor(0x7F, 0x00, 0x00),
}
SERIES_THEME = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3, MSO_THEME_COLOR.ACCENT_4]
SERIES_DASH = [MSO_LINE.SOLID, MSO_LINE.DASH, MSO_LINE.ROUND_DOT, MSO_LINE.LONG_DASH]
NON_CONTENT = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE,
               PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER)
CHART_FONT = Pt(12)
# A true minus sign: it matches slide text, and PowerPoint won't wrap a line after it.
PCT_LABEL = '0.0"%";\u22120.0"%"'
PCT_AXIS = '0"%";\u22120"%"'
MONEY_AXIS = "#,##0;\u2212#,##0"
GAP = Inches(0.3)
THEME_BODY_FONT = "+mn-lt"  # chart titles otherwise fall back to a fixed font


class TemplateError(ValueError):
    """The template lacks a layout or placeholder the deck needs."""


class Region:
    def __init__(self, left, top, width, height):
        self.left, self.top, self.width, self.height = int(left), int(top), int(width), int(height)

    def split_columns(self, count=2, gap=GAP):
        width = (self.width - gap * (count - 1)) // count
        return [Region(self.left + i * (width + gap), self.top, width, self.height) for i in range(count)]


def clear_slides(prs):
    """Remove any slides that ship with the template, keeping its master and layouts."""
    ids = prs.slides._sldIdLst
    for slide_id in list(ids):
        prs.part.drop_rel(slide_id.rId)
        ids.remove(slide_id)


def find_layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    available = ", ".join(repr(layout.name) for layout in prs.slide_layouts)
    raise TemplateError(
        f"Layout {name!r} is not in the template. Update 'layouts' in config.json. Available: {available}."
    )


def add_footer(slide, layout, text):
    """python-pptx doesn't copy footer or slide-number placeholders, so clone them from the layout."""
    for placeholder in layout.placeholders:
        kind = placeholder.placeholder_format.type
        if kind not in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
            continue
        element = copy.deepcopy(placeholder._element)
        element.xpath(".//p:cNvPr")[0].set("id", str(slide.shapes._next_shape_id))
        slide.shapes._spTree.append(element)
        if kind == PP_PLACEHOLDER.FOOTER:
            shape = slide.shapes[-1]
            shape.text_frame.text = text


def content_placeholders(slide):
    found = [p for p in slide.placeholders if p.placeholder_format.type not in NON_CONTENT]
    return sorted(found, key=lambda p: (p.left, p.top))


def take_region(placeholder):
    """Return a placeholder's geometry and remove it, so a chart or table can take its place."""
    region = Region(placeholder.left, placeholder.top, placeholder.width, placeholder.height)
    placeholder._element.getparent().remove(placeholder._element)
    return region


def free_area(prs, slide):
    """Space between the title and the footer on a layout with no content placeholder."""
    title = slide.shapes.title
    top = title.top + title.height + Inches(0.1)
    footers = [p.top for p in slide.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.FOOTER]
    bottom = min(footers) - Inches(0.1) if footers else int(prs.slide_height * 0.92)
    return Region(title.left, top, title.width, bottom - top)


def set_alt_text(shape, name, description):
    properties = shape._element.xpath(".//p:cNvPr")[0]
    properties.set("name", name)
    properties.set("descr", description)


def _inherited_size(placeholder, master_style):
    """Font size in points that a placeholder inherits from its layout, else from the master."""
    layout_sizes = placeholder._base_placeholder._element.xpath(".//a:lstStyle/a:lvl1pPr/a:defRPr/@sz")
    if layout_sizes:
        return int(layout_sizes[0]) / 100
    master = placeholder.part.slide.slide_layout.slide_master
    master_sizes = master._element.xpath(f"p:txStyles/p:{master_style}/a:lvl1pPr/a:defRPr/@sz")
    return int(master_sizes[0]) / 100 if master_sizes else 18


def fit_text(placeholder, master_style, minimum, space_after=0):
    """Shrink text that won't fit the placeholder, for templates with larger fonts or smaller boxes.

    PowerPoint only recalculates autofit when a person edits the text, so a generated
    deck has to set the size itself. The width is an estimate from the average character width.
    """
    size = _inherited_size(placeholder, master_style)
    width = max(placeholder.width - Inches(0.2), Inches(1)) / 914400
    available = placeholder.height / 914400
    paragraphs = placeholder.text_frame.paragraphs

    def needed(points):
        lines = sum(max(math.ceil(len(p.text) * 0.5 * points / 72 / width), 1) for p in paragraphs)
        return lines * points * 1.2 / 72 + len(paragraphs) * space_after / 72

    fitted = size
    while fitted > minimum and needed(fitted) > available:
        fitted -= 1
    if fitted < size:
        for paragraph in paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(fitted)


def write_title(title, text):
    title.text_frame.text = text
    fit_text(title, "titleStyle", minimum=20)


def write_bullets(placeholder, lines):
    frame = placeholder.text_frame
    frame.text = lines[0]
    for line in lines[1:]:
        frame.add_paragraph().text = line
    for paragraph in frame.paragraphs:
        paragraph.space_after = Pt(10)
    fit_text(placeholder, "bodyStyle", minimum=12, space_after=10)


def _muted(font):
    font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    font.color.brightness = 0.25


def _style_chart(chart, title, legend):
    chart.font.size = CHART_FONT
    _muted(chart.font)
    chart.has_title = True
    frame = chart.chart_title.text_frame
    frame.text = title
    run = frame.paragraphs[0].runs[0]
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.name = THEME_BODY_FONT
    run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    chart.chart_title.include_in_layout = False
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = CHART_FONT


def _style_axes(chart, number_format, bounds=None, major_unit=None):
    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
    value_axis.major_gridlines.format.line.width = Pt(0.75)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.number_format = number_format
    value_axis.tick_labels.number_format_is_linked = False
    if bounds:
        value_axis.minimum_scale, value_axis.maximum_scale = bounds
    if major_unit:
        value_axis.major_unit = major_unit
    category_axis = chart.category_axis
    category_axis.has_major_gridlines = False
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    category_axis.format.line.color.theme_color = MSO_THEME_COLOR.TEXT_1
    category_axis.format.line.color.brightness = 0.5


def add_column_chart(slide, region, categories, values, *, title, description, label_format=None,
                     axis_format=MONEY_AXIS, bounds=None, major_unit=None, highlight=None, tick_size=None):
    data = CategoryChartData()
    data.categories = categories
    data.add_series(title, values)
    frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, region.left, region.top, region.width,
                                   region.height, data)
    set_alt_text(frame, title, description)
    chart = frame.chart
    _style_chart(chart, title, legend=False)
    _style_axes(chart, axis_format, bounds, major_unit)
    if tick_size:
        chart.category_axis.tick_labels.font.size = tick_size
    plot = chart.plots[0]
    plot.gap_width = 55
    plot.vary_by_categories = False
    series = plot.series[0]
    series.invert_if_negative = False
    series.format.fill.solid()
    series.format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    if highlight is not None:
        point = series.points[highlight]
        point.format.fill.solid()
        point.format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
        # Without this a negative highlighted bar renders white.
        data_point = series._element.xpath(f'c:dPt[c:idx/@val="{highlight}"]')[0]
        data_point.insert(1, etree.Element(qn("c:invertIfNegative"), val="0"))
    if label_format:
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.number_format = label_format
        labels.number_format_is_linked = False
        labels.position = XL_LABEL_POSITION.OUTSIDE_END
        labels.show_value = True
        labels.font.size = CHART_FONT
        labels.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    return chart


def add_line_chart(slide, region, categories, series_map, *, title, description, axis_format=MONEY_AXIS,
                   label_format=None, tick_size=None):
    """Line chart with a legend when there are several series; labels the last point when asked."""
    data = CategoryChartData()
    data.categories = categories
    for name, values in series_map.items():
        data.add_series(name, values)
    frame = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, region.left, region.top, region.width,
                                   region.height, data)
    set_alt_text(frame, title, description)
    chart = frame.chart
    _style_chart(chart, title, legend=len(series_map) > 1)
    _style_axes(chart, axis_format)
    if tick_size:
        chart.category_axis.tick_labels.font.size = tick_size
    for index, series in enumerate(chart.plots[0].series):
        series.smooth = False
        series.format.line.width = Pt(2.25)
        series.format.line.color.theme_color = SERIES_THEME[index]
        series.format.line.dash_style = SERIES_DASH[index]
        series.marker.size = 8
        series.marker.format.fill.solid()
        series.marker.format.fill.fore_color.theme_color = SERIES_THEME[index]
        series.marker.format.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
        if label_format:
            values = list(series.values)
            label = series.points[len(values) - 1].data_label
            label.has_text_frame = True
            label.text_frame.text = label_format(values[-1])
            label.text_frame.paragraphs[0].runs[0].font.size = CHART_FONT
            label.text_frame.paragraphs[0].runs[0].font.color.theme_color = MSO_THEME_COLOR.TEXT_1
            label.position = XL_LABEL_POSITION.ABOVE
    return chart


def add_scatter_lines(slide, region, series_map, *, title, description, axis_format=MONEY_AXIS, x_title=None):
    """Lines on a numeric x axis. series_map = {name: [(x, y), ...]}."""
    data = XyChartData()
    for name, points in series_map.items():
        series = data.add_series(name)
        for x, y in points:
            series.add_data_point(x, y)
    frame = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, region.left, region.top,
                                   region.width, region.height, data)
    set_alt_text(frame, title, description)
    chart = frame.chart
    _style_chart(chart, title, legend=True)
    _style_axes(chart, axis_format)
    chart.category_axis.maximum_scale = max(x for points in series_map.values() for x, _ in points)
    chart.category_axis.minimum_scale = 0
    if x_title:
        axis_title = chart.category_axis.axis_title
        axis_title.text_frame.text = x_title
        run = axis_title.text_frame.paragraphs[0].runs[0]
        run.font.size = CHART_FONT
        run.font.bold = False
        run.font.name = THEME_BODY_FONT
    for index, series in enumerate(chart.plots[0].series):
        series.smooth = False
        series.format.line.width = Pt(2.25)
        series.format.line.color.theme_color = SERIES_THEME[index]
        series.format.line.dash_style = SERIES_DASH[index]
    return chart


def _bottom_border(cell):
    properties = cell._tc.get_or_add_tcPr()
    border = etree.SubElement(properties, qn("a:lnB"), w="9525", cap="flat", cmpd="sng", algn="ctr")
    fill = etree.SubElement(border, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr"), val="bg2")
    # Schema order: borders come before any fill in a:tcPr.
    properties.remove(border)
    properties.insert(0, border)


def _write_cell(cell, value, size, align, bold=False, header=False):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.08)
    cell.margin_top = cell.margin_bottom = Inches(0.03)
    paragraph = cell.text_frame.paragraphs[0]
    paragraph.alignment = align
    if value in STATUS_COLOURS and not header:
        dot = paragraph.add_run()
        dot.text = "● "
        dot.font.size = Pt(size)
        dot.font.color.rgb = STATUS_COLOURS[value]
    run = paragraph.add_run()
    run.text = str(value)
    run.font.size = Pt(size)
    run.font.bold = bold or header
    run.font.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1 if header else MSO_THEME_COLOR.TEXT_1
    if header:
        cell.fill.solid()
        cell.fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_2
    else:
        _bottom_border(cell)


def add_table(slide, region, header, rows, *, widths, name, size=14, row_height=0.42, numeric_from=1, bold_rows=(), left_columns=()):
    """Add a table. Columns from `numeric_from` onwards are right-aligned, except `left_columns`.

    Status cells get a coloured dot and a label.
    """
    count = len(rows) + 1
    height = min(int(Inches(row_height)) * count, region.height)
    frame = slide.shapes.add_table(count, len(header), region.left, region.top, region.width, height)
    set_alt_text(frame, name, f"{name}: {len(rows)} rows")
    table = frame.table
    # "No Style, No Grid": cell formatting below uses theme colours directly.
    frame._element.xpath(".//a:tableStyleId")[0].text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    table.horz_banding = False
    total = sum(widths)
    used = 0
    for index, column in enumerate(table.columns):
        column.width = region.width - used if index == len(widths) - 1 else int(region.width * widths[index] / total)
        used += column.width
    for row in table.rows:
        row.height = Emu(height // count)
    def alignment(column):
        return PP_ALIGN.RIGHT if column >= numeric_from and column not in left_columns else PP_ALIGN.LEFT

    for column, value in enumerate(header):
        align = alignment(column)
        _write_cell(table.cell(0, column), value, size, align, header=True)
    for row_index, row in enumerate(rows, start=1):
        for column, value in enumerate(row):
            align = alignment(column)
            _write_cell(table.cell(row_index, column), value, size, align, bold=row_index - 1 in bold_rows)
    return table


def add_kpi_card(slide, region, value, label, status=None):
    """Tinted stat card: large value, small label, optional status dot and text."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, region.left, region.top, region.width, region.height)
    card.adjustments[0] = 0.06
    card.name = f"KPI: {label}"
    card.fill.solid()
    card.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
    card.line.fill.background()
    card.shadow.inherit = False
    frame = card.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = Inches(0.2)

    first = frame.paragraphs[0]
    first.alignment = PP_ALIGN.LEFT
    run = first.add_run()
    run.text = value
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.theme_color = MSO_THEME_COLOR.TEXT_2

    second = frame.add_paragraph()
    second.alignment = PP_ALIGN.LEFT
    run = second.add_run()
    run.text = label
    run.font.size = Pt(13)
    run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1

    if status:
        third = frame.add_paragraph()
        third.alignment = PP_ALIGN.LEFT
        dot = third.add_run()
        dot.text = "● "
        dot.font.size = Pt(13)
        dot.font.color.rgb = STATUS_COLOURS[status]
        run = third.add_run()
        run.text = status
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    return card
