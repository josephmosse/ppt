# Automated IRRBB report

Turns one Excel workbook into a PowerPoint IRRBB report for ALCO. Python calculates the supervisory shocks,
ΔEVE, ΔNII and limit status, then builds native charts and tables you can still edit in PowerPoint. Slide titles
and commentary are written from the numbers by fixed rules, with no AI involved.

## Run it

```powershell
.\.venv\Scripts\python.exe build.py
```

The report is saved as `output\IRRBB_Report_<reporting date>.pptx`.

First-time setup on another machine:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Quarterly routine

Update `data\irrbb_input.xlsx`, then run the command above.

| Sheet | What to update |
|---|---|
| `Settings` | Institution, reporting date, currency, Tier 1, shock sizes |
| `Repricing_Gap` | Assets, liabilities and net derivatives by bucket, in millions, excluding own equity |
| `Curve` | Zero rates in % |
| `Limits` | Early warning, internal limit and SOT, in % of Tier 1 |
| `History` | Add last quarter's reported ΔEVE and ΔNII as a new row, in millions |
| `Assumptions` | Rows shown on the assumptions slide |
| `Actions` | Rows shown on the actions slide; `Done`, `Closed` or `Complete` counts as closed |

Keep the sheet and column names. If the workbook has a problem, the build stops and names the sheet and column.

## Use your own template

1. Save your company template as a `.pptx`. A `.potx` must be saved as `.pptx` first.
2. See its layout names:
   ```powershell
   .\.venv\Scripts\python.exe build.py --template templates\my_template.pptx --list-layouts
   ```
3. In `config.json`, set `template` and match each entry in `layouts` to one of your layout names:

| Key | The layout needs |
|---|---|
| `title` | A title and a subtitle |
| `content` | A title and one content placeholder |
| `two_content` | A title and two content placeholders: chart on the left, text on the right |
| `title_only` | A title |

Positions, fonts and colours then come from your template. Any slides already in it are removed from the output.
Status colours (green, amber, red) stay fixed, and are set in `STATUS_COLOURS` in `irrbb/pptx_helpers.py`.
Chart series use accent colours 1 to 4 of your theme, so check they stay distinguishable. Text is reduced when it
would overflow, using an estimate, so look through the first deck built on a new template.

## Change the content

| To change | Edit |
|---|---|
| Slide order, or add and remove slides | `DeckBuilder.build` in `irrbb/deck.py` |
| Titles and commentary rules | The slide methods in `irrbb/deck.py` |
| Calculations | `irrbb/calc.py` |
| Chart and table styling | `irrbb/pptx_helpers.py` |

## Example: how a slide is built (`eve_slide`)

`DeckBuilder.eve_slide` in `irrbb/deck.py` takes numbers that are already calculated and turns them into one slide,
in five steps. The values below come from the sample data.

**1. Get the data**

```python
i, r = self.inputs, self.results
worst = r.worst_eve            # "Parallel up"
value = r.eve_pct[worst]       # -10.27
```

`i` holds what was read from Excel, and `r` holds what `calc.py` worked out. The function does no risk calculation
itself; it only decides how to present the results.

**2. Create the slide and write the headline**

```python
slide, region, text = self._chart_and_text(f"{worst} is the most severe EVE scenario at ...")
```

This uses the template's `two_content` layout and returns three things:

- `slide`: the new slide
- `region`: the space on the left for the chart
- `text`: the placeholder on the right for the bullets

The headline here reads "Parallel up is the most severe EVE scenario at −10.3% of Tier 1, inside the 15% SOT". If the
loss passes 15%, "inside" becomes "above" automatically.

**3. Size the chart axis**

```python
sot = i.limits.loc["EVE", "SOT_Pct"]                # 15
bounds, unit = symmetric_axis(r.eve_pct, sot, 5)    # (-15, 15), 5
```

The axis always reaches at least the SOT, so readers can see how close each bar is to it. It also grows if the values
get larger; with much larger losses, for example, it went to ±60% in steps of 10.

**4. Draw the chart**

`ph.add_column_chart(...)` creates one column per scenario and highlights the worst one. The styling lives in
`pptx_helpers.py`, so `eve_slide` only says what to show.

**5. Write the commentary**

`ph.write_bullets(...)` turns each list item into a bullet:

- **Loss and change:** "Parallel up costs EUR 190m, 0.8 pp more severe than last quarter."
- **Headroom:** "Headroom to the 12% internal limit is 1.7 pp."
- **Best scenario:** "Parallel down is the most favourable scenario at +13.3% of Tier 1."

**How the data flows**

```text
Excel → data.py (read and check) → calc.py (ΔEVE, ΔNII) → eve_slide (choose the message) → pptx_helpers.py (draw it)
```

To change a sentence, edit the text in `eve_slide`. To change how the chart looks, edit `pptx_helpers.py`. To change
a result, edit `calc.py`.

## Calculation scope

This is a simplified, single-currency gap engine, and it won't tie exactly to a full ALM system:

- ΔEVE discounts each bucket's net gap as a zero-coupon flow at its midpoint, with continuous compounding.
- ΔNII covers buckets repricing within 12 months, on a constant balance sheet, for the two parallel scenarios.
- Shocks use the BCBS d368 functional form. Post-shock rates are floored at −150 bp at t = 0, rising 3 bp a year to 0%.
- Multi-currency aggregation, automatic options, and commercial margin treatment are not modelled.

Confirm the method, shock sizes and SOT thresholds against your own policy and current regulation before any
official use. The sample workbook and the assumptions in it are illustrative.

## Other commands

```powershell
.\.venv\Scripts\python.exe -m unittest discover tests   # calculation tests
.\.venv\Scripts\python.exe tools\make_sample_data.py    # OVERWRITES data\irrbb_input.xlsx with sample data
.\.venv\Scripts\python.exe tools\make_template.py       # recreates the starter template
```
