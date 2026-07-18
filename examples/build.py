"""Render examples/workbook.json into a .xlsx locally (no Open WebUI required).

    pip install openpyxl
    python examples/build.py
"""
import importlib.util
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

spec = json.loads((BASE / "examples" / "workbook.json").read_text())

_spec = importlib.util.spec_from_file_location(
    "generate_spreadsheets", BASE / "generate_spreadsheets.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

resolved = mod._resolve_template({**spec, "template": spec.get("template") or "financial"})
data, warnings = mod._build_workbook(resolved)
out = BASE / "examples" / "demo_budget.xlsx"
out.write_bytes(data)
print(f"OK: {out} ({len(data)} bytes)")
for w in warnings:
    print("  warn:", w)
