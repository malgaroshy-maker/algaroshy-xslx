# Formula Patterns Reference

Every derived value (totals, ratios, growth rates, percentages, differences) MUST be an Excel formula, never hardcoded.

## Aggregation Formulas

```python
ws['B10']  = '=SUM(B2:B9)'         # Sum range
ws['D20']  = '=AVERAGE(D2:D19)'    # Average
ws['E5']   = '=COUNT(B2:B9)'       # Count numbers
ws['F10']  = '=COUNTA(B2:B9)'      # Count non-empty
ws['G5']   = '=MAX(B2:B9)'         # Maximum
ws['H5']   = '=MIN(B2:B9)'         # Minimum
```

## Conditional Aggregation

```python
ws['C10']  = '=SUMIF(A2:A9,"West",B2:B9)'       # Sum with condition
ws['D10']  = '=SUMIFS(D2:D9,A2:A9,"West",B2:B9,">100")'  # Sum with multiple conditions
ws['E10']  = '=COUNTIF(A2:A9,"Active")'          # Count if
ws['F10']  = '=AVERAGEIF(A2:A9,"Yes",B2:B9)'    # Average if
```

## Growth & Change

```python
ws['C5']   = '=(B5-B4)/B4*100'     # Period-over-period growth %
ws['D5']   = '=C5-B5'              # Simple difference
ws['E5']   = '=C5/B5-1'            # Growth ratio
```

## Percentage of Total

```python
ws['F5']   = '=B5/B$10*100'        # % of total (absolute row ref)
ws['G5']   = '=B5/SUM(B$2:B$9)*100' # % of dynamic total
```

## Running Total / Cumulative

```python
ws['H5']   = '=H4+F5'              # Running sum (first row hardcoded)
# For row 2 (first data): ws['H2'] = '=F2'
```

## Safety Formulas

```python
ws['D5']   = '=IF(B5>0,C5/B5*100,0)'     # Avoid #DIV/0!
ws['E5']   = '=IFERROR(VLOOKUP(A5,Data!A:B,2,FALSE),"N/A")'  # Graceful lookup miss
ws['F5']   = '=IF(A5="","",B5+C5)'        # Skip empty rows
```

## Row-Relative Pattern

Use Python f-strings with the row number `{rn}` to generate formulas:

```python
for i, row in df.iterrows():
    rn = i + 2  # Excel row (1-indexed, header at row 1)
    ws.cell(row=rn, column=3, value=f'=(B{rn}-B{rn-1})/B{rn-1}*100')  # growth
    ws.cell(row=rn, column=4, value=f'=C{rn}-B{rn}')                   # variance
    ws.cell(row=rn, column=5, value=f'=IF(B{rn}>0,D{rn}/B{rn}*100,0)') # %
```

## Cross-Sheet References

```python
ws.cell(row=rn, column=6, value=f"=SUMIF('Sales Data'!A:A,A{rn},'Sales Data'!B:B)")
ws.cell(row=rn, column=7, value=f"='Revenue Trends'!B{rn}")
# Cross-sheet formulas use GREEN font (color='00B050')
```

## Financial Colors

| Cell Role | Font Color | Code |
|-----------|-----------|------|
| Hard-coded input/assumption | Blue | `0000FF` |
| Formula / computed result | Black | `000000` |
| Cross-sheet reference | Green | `00B050` |

```python
input_font = Font(name='Arial', size=10, color='0000FF')
formula_font = Font(name='Arial', size=10, color='000000')
cross_font = Font(name='Arial', size=10, color='00B050')
```
