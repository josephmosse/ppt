"""Build the IRRBB report. Every title and bullet is derived from the numbers."""
import math

import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt

from . import pptx_helpers as ph
from .calc import (EARLY_WARNING, LIMIT_BREACH, NII_SCENARIOS, SCENARIOS, SOT_BREACH, STATUS_ORDER, WITHIN,
                   shock_bp, status_for)
from .data import Inputs

CLOSED_STATUSES = {"done", "closed", "complete", "completed"}
SHOCK_TENORS = [("3M", 0.25), ("1Y", 1), ("5Y", 5), ("10Y", 10), ("20Y", 20)]


def minus(text):
    """Use a true minus sign: it reads better and PowerPoint will not wrap a line after it."""
    return text.replace("-", "\u2212")


def pct(value):
    return minus(f"{value:.1f}%")


def signed_pct(value):
    return minus(f"{value:+.1f}%")


def signed_money(value):
    return minus(f"{value:+,.0f}")


def quarter(date):
    return f"Q{date.quarter} {date.year}"


def worst_loss(series):
    return max(-series.min(), 0.0)


def change_phrase(current, prior):
    """Compare the size of the worst loss, in percentage points of Tier 1."""
    delta = worst_loss(current) - worst_loss(prior)
    if abs(delta) < 0.05:
        return "unchanged on last quarter"
    direction = "more severe" if delta > 0 else "less severe"
    return f"{abs(delta):.1f} pp {direction} than last quarter"


def metric_status(inputs, metric, value):
    row = inputs.limits.loc[metric]
    return status_for(value, row["Early_Warning_Pct"], row["Limit_Pct"], row["SOT_Pct"])


def sot_phrase(inputs, metric, value):
    sot = inputs.limits.loc[metric, "SOT_Pct"]
    relation = "above" if -value > sot else "inside"
    return f"{relation} the {sot:g}% SOT"


def headroom_phrase(inputs, metric, value):
    limit = inputs.limits.loc[metric, "Limit_Pct"]
    headroom = limit - max(-value, 0)
    if headroom < 0:
        return f"Exceeds the {limit:g}% internal limit by {-headroom:.1f} pp."
    return f"Headroom to the {limit:g}% internal limit is {headroom:.1f} pp."


def overall_headline(eve_status, nii_status):
    statuses = {"ΔEVE": eve_status, "ΔNII": nii_status}
    worst = max(statuses.values(), key=STATUS_ORDER.index)
    names = " and ".join(name for name, status in statuses.items() if status == worst)
    if worst == SOT_BREACH:
        return f"IRRBB breaches the supervisory outlier test on {names}"
    if worst == LIMIT_BREACH:
        return f"IRRBB breaches the internal limit on {names}, with no SOT breach"
    if worst == EARLY_WARNING:
        return f"IRRBB stays within limits, with an early warning on {names}"
    return "IRRBB stays within all internal limits and supervisory outlier tests"


def open_actions(inputs):
    actions = inputs.actions
    return actions[~actions["Status"].astype(str).str.strip().str.lower().isin(CLOSED_STATUSES)]


def symmetric_axis(values, threshold, step):
    """Symmetric bounds and tick step reaching at least the SOT, so distance to it is visible at a glance."""
    reach = max(threshold, max(abs(v) for v in values) * 1.1)
    while reach / step > 6:  # keep large values from crowding the axis with ticks
        step *= 2
    bound = math.ceil(reach / step) * step
    return (-bound, bound), step


class DeckBuilder:
    def __init__(self, inputs: Inputs, results, config):
        self.inputs, self.results = inputs, results
        self.layouts = config["layouts"]
        self.prs = Presentation(config["template"])
        ph.clear_slides(self.prs)
        for key in self.layouts:  # fail before any slide is built
            ph.find_layout(self.prs, self.layouts[key])
        self.footer = f"{inputs.institution} | IRRBB report | {inputs.reporting_date:%d %b %Y}"

    def _slide(self, layout_key, title):
        layout = ph.find_layout(self.prs, self.layouts[layout_key])
        slide = self.prs.slides.add_slide(layout)
        if slide.shapes.title is None:
            raise ph.TemplateError(f"Layout {layout.name!r} has no title placeholder.")
        ph.write_title(slide.shapes.title, title)
        if layout_key != "title":
            ph.add_footer(slide, layout, self.footer)
        return slide

    def _chart_and_text(self, title):
        """Two-content slide: returns the slide, the chart's region and the text placeholder."""
        slide = self._slide("two_content", title)
        placeholders = ph.content_placeholders(slide)
        if len(placeholders) < 2:
            raise ph.TemplateError(f"Layout {self.layouts['two_content']!r} needs two content placeholders.")
        return slide, ph.take_region(placeholders[0]), placeholders[1]

    def _full_region(self, slide):
        placeholders = ph.content_placeholders(slide)
        if not placeholders:
            raise ph.TemplateError(f"Layout {self.layouts['content']!r} needs a content placeholder.")
        return ph.take_region(placeholders[0])

    def title_slide(self):
        i = self.inputs
        slide = self._slide("title", "Interest Rate Risk in the Banking Book")
        subtitle = next((p for p in slide.placeholders if p.placeholder_format.idx == 1), None)
        if subtitle is not None:
            subtitle.text_frame.text = f"{i.institution} | {i.committee}"
            subtitle.text_frame.add_paragraph().text = (
                f"Position at {i.reporting_date:%d %B %Y} | Prepared by {i.prepared_by}"
            )

    def executive_summary(self):
        i, r = self.inputs, self.results
        eve, nii = r.eve_pct.min(), r.nii_pct.min()
        eve_status, nii_status = metric_status(i, "EVE", eve), metric_status(i, "NII", nii)
        slide = self._slide("content", overall_headline(eve_status, nii_status))

        body = ph.content_placeholders(slide)[0]
        left, top, width, height = body.left, body.top, body.width, body.height
        card_height = Inches(1.75)
        cards = ph.Region(left, top, width, card_height).split_columns(3)
        ph.add_kpi_card(slide, cards[0], pct(eve), f"Worst ΔEVE, % of Tier 1 ({r.worst_eve})", eve_status)
        ph.add_kpi_card(slide, cards[1], pct(nii), f"Worst ΔNII, % of Tier 1 ({r.worst_nii})", nii_status)
        ph.add_kpi_card(slide, cards[2], f"{i.currency} {i.tier1:,.0f}m", f"Tier 1 capital at {i.reporting_date:%d %b %Y}")

        offset = card_height + ph.GAP
        body.left, body.top, body.width, body.height = left, top + offset, width, height - offset

        largest = r.gap.loc[r.gap["Net_Gap"].abs().idxmax()]
        pending = open_actions(i)
        lines = [
            f"ΔEVE: {r.worst_eve} costs {i.currency} {-r.eve.min():,.0f}m, {change_phrase(r.eve_pct, r.prior_eve_pct)}.",
            f"ΔNII: {r.worst_nii} changes 12-month NII by {i.currency} {signed_money(r.nii.min())}m, "
            f"{change_phrase(r.nii_pct, r.prior_nii_pct)}.",
            f"Largest repricing gap: {i.currency} {signed_money(largest['Net_Gap'])}m in the {largest['Bucket']} bucket.",
            f"{len(pending)} of {len(i.actions)} management actions remain open.",
        ]
        ph.write_bullets(body, lines)

    def limit_monitoring(self):
        i, r = self.inputs, self.results
        statuses = [metric_status(i, "EVE", r.eve_pct[s]) for s in SCENARIOS]
        statuses += [metric_status(i, "NII", r.nii_pct[s]) for s in NII_SCENARIOS]
        flagged = sum(status != WITHIN for status in statuses)
        if flagged:
            title = f"{flagged} of {len(statuses)} scenario measures {'sits' if flagged == 1 else 'sit'} above early warning"
        else:
            title = f"All {len(statuses)} scenario measures are below early warning"
        slide = self._slide("content", title)
        region = self._full_region(slide)

        def row(metric, label, scenario, current, prior):
            limits = i.limits.loc[metric]
            return [
                f"{label} {scenario}", pct(current[scenario]), pct(prior[scenario]),
                f"{limits['Early_Warning_Pct']:g}%", f"{limits['Limit_Pct']:g}%", f"{limits['SOT_Pct']:g}%",
                metric_status(i, metric, current[scenario]),
            ]

        rows = [row("EVE", "ΔEVE", s, r.eve_pct, r.prior_eve_pct) for s in SCENARIOS]
        rows += [row("NII", "ΔNII", s, r.nii_pct, r.prior_nii_pct) for s in NII_SCENARIOS]
        header = ["Measure, % of Tier 1", quarter(i.reporting_date), quarter(r.prior_date),
                  "Early warning", "Limit", "SOT", "Status"]
        ph.add_table(slide, region, header, rows, widths=[3.4, 1.2, 1.2, 1.3, 0.9, 0.9, 2.0],
                     name="Limit monitoring by scenario", size=14, row_height=0.5, left_columns=(6,))

    def eve_slide(self):
        """Build the ΔEVE slide: a headline, a chart of the six scenarios, and three commentary bullets."""
        # i = what was read from Excel (currency, limits...); r = what calc.py worked out (ΔEVE, ΔNII...).
        i, r = self.inputs, self.results

        # Find the scenario with the lowest ΔEVE, e.g. worst = "Parallel up" and value = -10.27 (% of Tier 1).
        worst = r.worst_eve
        value = r.eve_pct[worst]

        # Create the slide from the template's two-content layout, with the headline written from the numbers.
        # Returns the slide, the space on the left for the chart, and the text placeholder on the right.
        slide, region, text = self._chart_and_text(
            f"{worst} is the most severe EVE scenario at {pct(value)} of Tier 1, {sot_phrase(i, 'EVE', value)}"
        )

        # Read the SOT threshold from the Limits sheet (15 in the sample), then size the value axis.
        # The axis always reaches at least the SOT, so the distance to it is visible, and 5 is the starting tick step.
        sot = i.limits.loc["EVE", "SOT_Pct"]
        bounds, unit = symmetric_axis(r.eve_pct, sot, 5)

        # Draw a native PowerPoint column chart, with one column per scenario.
        ph.add_column_chart(
            slide, region, SCENARIOS, [round(r.eve_pct[s], 2) for s in SCENARIOS],  # column names and heights
            title=f"ΔEVE, % of Tier 1 (SOT: {sot:g}% loss)",  # :g drops trailing zeros, so 15.0 shows as 15
            description="; ".join(f"{s}: {pct(r.eve_pct[s])}" for s in SCENARIOS),  # alt text for screen readers
            label_format=ph.PCT_LABEL, axis_format=ph.PCT_AXIS,  # number formats for bar labels and the axis
            bounds=bounds, major_unit=unit,  # axis range and tick step worked out above
            highlight=SCENARIOS.index(worst), tick_size=Pt(11),  # colour the worst column differently
        )

        # Find the scenario with the highest ΔEVE, for the last bullet.
        best = r.eve_pct.idxmax()

        # Fill the right-hand placeholder. Each list item becomes one bullet, written from the numbers.
        # Bullet 1: the loss in money and the change on last quarter. -r.eve[worst] shows the loss as a positive amount.
        # Bullet 2: the headroom to the internal limit, or how far it is exceeded.
        # Bullet 3: the most favourable scenario.
        ph.write_bullets(text, [
            f"{worst} costs {i.currency} {-r.eve[worst]:,.0f}m, {change_phrase(r.eve_pct, r.prior_eve_pct)}.",
            headroom_phrase(i, "EVE", value),
            f"{best} is the most favourable scenario at {signed_pct(r.eve_pct[best])} of Tier 1.",
        ])

    def nii_slide(self):
        i, r = self.inputs, self.results
        worst = r.worst_nii
        value = r.nii_pct[worst]
        slide, region, text = self._chart_and_text(
            f"{worst} is the weaker NII scenario at {pct(value)} of Tier 1, {sot_phrase(i, 'NII', value)}"
        )
        sot = i.limits.loc["NII", "SOT_Pct"]
        bounds, unit = symmetric_axis(r.nii_pct, sot, 1)
        ph.add_column_chart(
            slide, region, NII_SCENARIOS, [round(r.nii_pct[s], 2) for s in NII_SCENARIOS],
            title=f"12-month ΔNII, % of Tier 1 (SOT: {sot:g}% loss)",
            description="; ".join(f"{s}: {pct(r.nii_pct[s])}" for s in NII_SCENARIOS),
            label_format=ph.PCT_LABEL, axis_format=ph.PCT_AXIS,
            bounds=bounds, major_unit=unit,
            highlight=NII_SCENARIOS.index(worst),
        )
        ph.write_bullets(text, [
            f"{worst} changes 12-month NII by {i.currency} {signed_money(r.nii[worst])}m, "
            f"{change_phrase(r.nii_pct, r.prior_nii_pct)}.",
            headroom_phrase(i, "NII", value),
            "Measured on a constant balance sheet with an instantaneous shock.",
        ])

    def gap_slide(self):
        i, r = self.inputs, self.results
        gap = r.gap
        largest = gap.loc[gap["Net_Gap"].abs().idxmax()]
        within_year = gap.loc[gap["Midpoint_Years"] < 1, "Net_Gap"].sum()
        slide = self._slide(
            "title_only",
            f"The largest repricing gap is {i.currency} {signed_money(largest['Net_Gap'])}m in {largest['Bucket']}; "
            f"the gap repricing within 1Y is {i.currency} {signed_money(within_year)}m",
        )
        left, right = ph.free_area(self.prs, slide).split_columns(2)
        buckets = gap["Bucket"].tolist()
        ph.add_column_chart(
            slide, left, buckets, gap["Net_Gap"].round(1).tolist(),
            title=f"Net repricing gap by bucket, {i.currency} m",
            description=f"Largest gap {largest['Net_Gap']:,.0f} in {largest['Bucket']}",
            tick_size=Pt(10),
        )
        ph.add_line_chart(
            slide, right, buckets, {"Cumulative gap": gap["Cumulative_Gap"].round(1).tolist()},
            title=f"Cumulative repricing gap, {i.currency} m",
            description=f"Cumulative gap ends at {gap['Cumulative_Gap'].iloc[-1]:,.0f}",
            label_format=lambda v: minus(f"{v:,.0f}"), tick_size=Pt(10),
        )

    def trend_slide(self):
        r = self.results
        trend = r.trend
        count = len(trend)
        slide, region, text = self._chart_and_text(
            f"Worst ΔEVE moved from {pct(trend['Worst_EVE_Pct'].iloc[0])} to {pct(trend['Worst_EVE_Pct'].iloc[-1])} "
            f"of Tier 1 over {count} quarters"
        )
        ph.add_line_chart(
            slide, region, [quarter(d) for d in trend["Date"]],
            {"Worst ΔEVE": trend["Worst_EVE_Pct"].round(2).tolist(), "Worst ΔNII": trend["Worst_NII_Pct"].round(2).tolist()},
            title="Worst-scenario change, % of Tier 1",
            description=f"{count} quarters of worst-scenario ΔEVE and ΔNII",
            axis_format=ph.PCT_AXIS, label_format=pct, tick_size=Pt(11),
        )
        peak = trend.loc[trend["Worst_EVE_Pct"].idxmin()]
        ph.write_bullets(text, [
            f"The most severe ΔEVE in the period was {pct(peak['Worst_EVE_Pct'])}, in {quarter(peak['Date'])}.",
            f"Worst ΔNII moved from {pct(trend['Worst_NII_Pct'].iloc[0])} to {pct(trend['Worst_NII_Pct'].iloc[-1])}.",
            "Earlier quarters come from the History sheet; the latest is calculated in this run.",
        ])

    def pillar3_slide(self):
        i, r = self.inputs, self.results
        slide = self._slide("content", f"Pillar 3 view: ΔEVE and ΔNII by supervisory scenario, {i.currency} m")
        region = self._full_region(slide)
        current, prior = quarter(i.reporting_date), quarter(r.prior_date)
        rows = []
        for scenario in SCENARIOS:
            has_nii = scenario in NII_SCENARIOS
            rows.append([
                scenario, signed_money(r.eve[scenario]), signed_money(r.prior_eve[scenario]),
                signed_money(r.nii[scenario]) if has_nii else "n/a",
                signed_money(r.prior_nii[scenario]) if has_nii else "n/a",
            ])
        rows.append(["Tier 1 capital", f"{i.tier1:,.0f}", f"{r.prior_tier1:,.0f}", f"{i.tier1:,.0f}", f"{r.prior_tier1:,.0f}"])
        header = ["Supervisory scenario", f"ΔEVE {current}", f"ΔEVE {prior}", f"ΔNII {current}", f"ΔNII {prior}"]
        ph.add_table(slide, region, header, rows, widths=[3.2, 2, 2, 2, 2], name="Pillar 3 IRRBB table",
                     size=14, row_height=0.52, bold_rows=(len(rows) - 1,))

    def shock_slide(self):
        i = self.inputs
        slide, chart_region, text = self._chart_and_text(
            f"Supervisory {i.currency} shocks: ±{i.parallel_bp:g} bp parallel, ±{i.short_bp:g} bp short end, "
            f"±{i.long_bp:g} bp long end"
        )
        table_region = ph.take_region(text)
        grid = np.linspace(0, 25, 51)
        plotted = ["Parallel up", "Short up", "Steepener", "Flattener"]
        params = (i.parallel_bp, i.short_bp, i.long_bp)
        ph.add_scatter_lines(
            slide, chart_region,
            {s: [(float(t), round(float(shock_bp(s, t, *params)), 1)) for t in grid] for s in plotted},
            title="Rate shock by maturity, bp (down shocks mirror the up shocks; before floor)",
            description="Shock in basis points by maturity for parallel, short, steepener and flattener scenarios",
            x_title="Maturity, years",
        )
        rows = [
            [s, *[minus(f"{float(shock_bp(s, t, *params)):+.0f}") for _, t in SHOCK_TENORS]]
            for s in SCENARIOS
        ]
        ph.add_table(slide, table_region, ["Shock, bp", *[label for label, _ in SHOCK_TENORS]], rows,
                     widths=[2.3, 1, 1, 1, 1, 1], name="Shock by tenor", size=12, row_height=0.42)

    def assumptions_slide(self):
        i = self.inputs
        slide = self._slide("content", "The results rest on these modelling assumptions")
        region = self._full_region(slide)
        rows = [[str(row["Item"]), str(row["Assumption"])] for _, row in i.assumptions.iterrows()]
        ph.add_table(slide, region, ["Item", "Assumption"], rows, widths=[3, 8.5], name="Modelling assumptions",
                     size=14, row_height=0.6, numeric_from=2)

    def actions_slide(self):
        i = self.inputs
        pending = open_actions(i)
        dated = pending["Due"].dropna()
        title = f"{len(pending)} management actions remain open"
        if not dated.empty:
            title += f"; the next is due {dated.min():%d %b %Y}"
        slide = self._slide("content", title)
        region = self._full_region(slide)
        rows = [
            [str(row["Action"]), str(row["Owner"]), f"{row['Due']:%d %b %Y}" if row["Due"] == row["Due"] else "", str(row["Status"])]
            for _, row in i.actions.iterrows()
        ]
        ph.add_table(slide, region, ["Action", "Owner", "Due", "Status"], rows, widths=[6.4, 2, 1.5, 1.6],
                     name="Management actions", size=14, row_height=0.6, numeric_from=4)

    def build(self, output):
        i = self.inputs
        self.prs.core_properties.title = f"IRRBB report {i.reporting_date:%Y-%m-%d}"
        self.prs.core_properties.author = i.prepared_by
        self.title_slide()
        self.executive_summary()
        self.limit_monitoring()
        self.eve_slide()
        self.nii_slide()
        self.gap_slide()
        self.trend_slide()
        self.pillar3_slide()
        self.shock_slide()
        self.assumptions_slide()
        self.actions_slide()
        self.prs.save(output)
        return len(self.prs.slides)
