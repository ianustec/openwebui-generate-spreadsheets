# Assets

Screenshots used by the README. All are from the example workbook
[`examples/workbook.json`](../examples/workbook.json) → [`examples/demo_budget.xlsx`](../examples/demo_budget.xlsx).

- `inputs.png` — Assumptions sheet (yellow input cells)
- `table.png` — P&L with Excel Table, formulas, data bars
- `matrix.png` — Revenue by region
- `chart.png` — KPI row + native bar chart
- `hero.png` — 3-panel preview banner

Regenerate:

```bash
python examples/build.py
soffice --headless --convert-to pdf examples/demo_budget.xlsx
pdftoppm -png -r 120 demo_budget.pdf page
```
