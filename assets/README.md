# Assets

Screenshots from the English demo workbook
[`examples/workbook.json`](../examples/workbook.json) → [`examples/demo_budget.xlsx`](../examples/demo_budget.xlsx)
(**Northwind FY2026 Operating Plan** — Assumptions, P&L, By segment, Pipeline, Dashboard).

| File | Sheet |
|---|---|
| `inputs.png` | Assumptions (yellow editable inputs) |
| `table.png` | Consolidated P&L with Excel Table + data bars |
| `matrix.png` | Revenue by segment × quarter (color scale) |
| `pipeline.png` | Sales pipeline with weighted formulas + validation |
| `chart.png` | Dashboard KPIs + bar + doughnut charts |
| `hero.png` | 3-panel README banner |

Regenerate:

```bash
python examples/build.py
# optional: set landscape print areas, then:
soffice --headless --convert-to pdf examples/demo_budget.xlsx
pdftoppm -png -r 160 demo_budget.pdf page
# trim + shadow each page into assets/*.png
```
