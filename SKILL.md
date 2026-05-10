---
name: algaroshy-xlsx
description: Ultimate Excel skill merging the best of Anthropic and MiniMax approaches. Open, create, read, analyze, edit, or validate Excel/spreadsheet files (.xlsx, .xlsm, .csv, .tsv). Formula-first philosophy (every derived value is a live Excel formula, never hardcoded). Native openpyxl charts (not PNG images). Financial color coding (blue=hardcoded input, black=formula, green=cross-sheet reference). XML unpack→edit→pack for existing files (zero format loss). Includes formula validation and audit scripts.
triggers: Excel, xlsx, spreadsheet, csv, pivot table, financial model, formula, data analysis, workbook, analysis
license: MIT
metadata:
  version: 1.0.0
  category: productivity
  author: Algaroshy
  sources:
    - Anthropic xlsx skill (Apache 2.0) — openpyxl, pandas patterns
    - MiniMax minimax-xlsx (MIT) — formula-first philosophy, financial colors, formula_check.py
---

# Algaroshy XLSX — The Merged Excel Skill

Formula-first, chart-native, audit-ready. Combines openpyxl reliability with MiniMax formula philosophy.

## Quick Reference

| Task | Method | Key Rule |
|------|--------|----------|
| **READ** | `pandas` | Discovery first, then analysis |
| **CREATE** | `openpyxl` + formulas | Every derived value = Excel formula, never hardcoded |
| **ADD CHART** | `openpyxl.chart` | Native Excel chart objects, not PNG images |
| **EDIT** | XML unpack→edit→pack | Preserve all formatting, only touch target cells |
| **VALIDATE** | `scripts/formula_check.py` | Zero formula errors before delivery |
| **STYLE** | `scripts/style_apply.py` | Financial colors + conditional fills |

## Formula-First Philosophy (MANDATORY)

**Every calculated cell MUST be an Excel formula, not a hardcoded number from Python.**

### WRONG — Hardcoding Calculated Values
```python
total = df['Sales'].sum()          # Bad: computing in Python
ws['B10'] = total                   # Bad: hardcoding 50000

growth = (df['Revenue'].iloc[-1] - df['Revenue'].iloc[0]) / df['Revenue'].iloc[0]
ws['C5'] = growth                   # Bad: hardcoded 0.15

avg = sum(values) / len(values)
ws['D20'] = avg                     # Bad: hardcoded 42.5
```

### CORRECT — Using Excel Formulas
```python
ws['B10'] = '=SUM(B2:B9)'           # Good: Excel calculates
ws['C5'] = '=(C4-C2)/C2'           # Good: growth rate as formula
ws['D20'] = '=AVERAGE(D2:D19)'     # Good: average as formula
ws['E5'] = '=SUMIF(A2:A9,"West",B2:B9)'  # Good: conditional sum
```

**Exception:** Input data cells (raw values read from source) CAN be hardcoded. Only DERIVED values (totals, ratios, percentages, growth rates, differences) MUST be formulas.

## Financial Color Standard

| Cell Role | Font Color | Hex Code | When |
|-----------|-----------|----------|------|
| Hard-coded input / assumption | Blue | `0000FF` (openpyxl: `Colors.BLUE`) | Raw data cells |
| Formula / computed result | Black | `000000` (default) | All `=SUM()`, `=AVERAGE()`, etc. |
| Cross-sheet reference | Green | `00B050` | Links to other sheets |
| Key assumption cell | Yellow bg | `FFFF00` | Cells meant to be changed |

## Conditional Fills (for data visualization)

| Condition | Fill | Hex |
|-----------|------|-----|
| Positive / Good | Green | `C6EFCE` |
| Warning / Medium | Yellow | `FFEB9C` |
| Negative / Bad | Red | `FFC7CE` |

## CREATE — New Workbook from Scratch

Use **openpyxl** for reliable creation. Always use formulas.

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
import pandas as pd

# Read source
df = pd.read_excel('source.xlsx')

# Create workbook
wb = Workbook()
ws = wb.active
ws.title = 'Analysis'

# Headers
headers = ['Month', 'Revenue', 'Growth %']
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=c, value=h)
    cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='2F5496')

# Data rows with formula for growth
for i, (_, row) in enumerate(df.iterrows()):
    r = i + 2
    ws.cell(row=r, column=1, value=row['Month'])
    ws.cell(row=r, column=2, value=int(row['Revenue']))
    if i > 0:
        ws.cell(row=r, column=3, value=f'=(B{r}-B{r-1})/B{r-1}*100')
    else:
        ws.cell(row=r, column=3, value='N/A')

# TOTAL row with formula
total_row = len(df) + 2
ws.cell(row=total_row, column=1, value='TOTAL').font = Font(bold=True)
ws.cell(row=total_row, column=2, value=f'=SUM(B2:B{total_row-1})')

# Add chart
chart = BarChart()
chart.title = 'Revenue by Month'
chart.x_axis.title = 'Month'
chart.y_axis.title = 'Revenue ($)'
data_ref = Reference(ws, min_col=2, min_row=1, max_row=total_row-1)
cats_ref = Reference(ws, min_col=1, min_row=2, max_row=total_row-1)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
chart.width = 18
chart.height = 10
ws.add_chart(chart, 'E3')

wb.save('output.xlsx')
```

## READ — Analyze Existing Data

Use `pandas` for data analysis. Never modify the source file during READ.

```python
import pandas as pd

# Read Excel
df = pd.read_excel('file.xlsx')
all_sheets = pd.read_excel('file.xlsx', sheet_name=None)

# Analyze
df.head()
df.info()
df.describe()

# Group and aggregate
monthly = df.groupby(df['Date'].dt.to_period('M')).agg({'Revenue': 'sum'})
```

## EDIT — Modify Existing File with Zero Format Loss

Never use openpyxl round-trip on existing files (corrupts VBA, pivots, sparklines). Use XML unpack→edit→pack.

```bash
python3 SKILL_DIR/scripts/xlsx_unpack.py input.xlsx /tmp/xlsx_work/
# Edit XML files in /tmp/xlsx_work/ using the Edit tool
python3 SKILL_DIR/scripts/xlsx_pack.py /tmp/xlsx_work/ output.xlsx
```

## VALIDATE — Formula Check (MANDATORY before delivery)

Run after every CREATE or EDIT:

```bash
python3 SKILL_DIR/scripts/formula_check.py output.xlsx --json
```

Exit code 0 = safe. If errors found, fix and re-run.

## Chart Best Practices

```python
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint

# Bar chart
chart = BarChart()
chart.type = 'col'
chart.title = 'Revenue by Region'
chart.y_axis.title = 'Revenue ($)'
chart.x_axis.title = 'Region'
chart.style = 10
chart.width = 18
chart.height = 12

# Pie chart
pie = PieChart()
pie.title = 'Revenue Distribution'
pie.width = 14
pie.height = 14
data = Reference(ws, min_col=2, min_row=1, max_row=5)
cats = Reference(ws, min_col=1, min_row=2, max_row=5)
pie.add_data(data, titles_from_data=True)
pie.set_categories(cats)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showCatName = True

# Line chart with markers
line = LineChart()
line.title = 'Monthly Trend'
line.y_axis.title = 'Revenue ($)'
line.marker.symbol = 'circle'
line.marker.size = 6

ws.add_chart(chart, 'F3')
```

## Column Width & Row Height

```python
# Auto-width based on content
for col in ws.columns:
    max_len = max((len(str(cell.value or '')) for cell in col), default=10)
    ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 45)

# Set header row height
ws.row_dimensions[1].height = 28
```

## Number Formatting

```python
ws.cell(row=r, column=c).number_format = '#,##0'        # Currency
ws.cell(row=r, column=c).number_format = '0.0%'         # Percentage
ws.cell(row=r, column=c).number_format = '0.0"%"'       # Percentage with sign
ws.cell(row=r, column=c).number_format = '0.0"x"'       # Multiples
ws.cell(row=r, column=c).number_format = '#,##0.00'     # 2 decimals
ws.cell(row=r, column=c).number_format = '#,##0;($#,##0);-'  # Negatives in parens
```

## Utility Scripts

```bash
# Formula validation
python3 SKILL_DIR/scripts/formula_check.py file.xlsx --json

# Style audit - check formula vs hardcode ratio  
python3 SKILL_DIR/scripts/formula_audit.py file.xlsx

# Apply financial colors to openpyxl workbook
python3 SKILL_DIR/scripts/style_apply.py file.xlsx --output styled.xlsx

# XML unpack/pack for editing existing files
python3 SKILL_DIR/scripts/xlsx_unpack.py input.xlsx /tmp/work/
python3 SKILL_DIR/scripts/xlsx_pack.py /tmp/work/ output.xlsx
```
