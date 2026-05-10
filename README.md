# Algaroshy XLSX | الجروشي إكسلسكس

> The ultimate Excel agent skill — formula-first, chart-native, audit-ready.
> 
> مهارة إكسل متكاملة للوكيل الذكي — صيغ حية، رسوم بيانية أصلية، جاهزة للتدقيق.

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Skills](https://img.shields.io/badge/skills.sh-algaroshy--xlsx-green)](https://skills.sh)
[![OpenCode](https://img.shields.io/badge/OpenCode-compatible-purple)](https://opencode.ai)
[![Claude Code](https://img.shields.io/badge/Claude%20Code-compatible-orange)](https://claude.ai/code)
[![Codex](https://img.shields.io/badge/Codex-compatible-black)](https://github.com/openai/codex)

Merges the best of [Anthropic's xlsx skill](https://github.com/anthropics/skills) and [MiniMax's minimax-xlsx](https://github.com/MiniMax-AI/skills) into one production-grade skill.

يدمج أفضل ما في مهارتي Anthropic و MiniMax في مهارة واحدة جاهزة للإنتاج.

## One-Liner Install

Just paste this into your OpenCode agent:

> install this excel skill in opencode: https://github.com/malgaroshy-maker/algaroshy-xslx.git

Or in Arabic:

> ثبت مهارة الإكسل هذه في opencode: https://github.com/malgaroshy-maker/algaroshy-xslx.git

## Why Algaroshy XLSX? | لماذا الجروشي إكسلسكس؟

| | Anthropic xlsx | MiniMax xlsx | **Algaroshy xlsx** |
|---|---|---|---|
| **Engine** | openpyxl (hardcoded) | Raw OOXML (fragile) | **openpyxl + formulas** |
| **Formulas** | 0 (all hardcoded) | 53 (broken XML) | **209 live formulas** |
| **Charts** | PNG images | None | **Native Excel charts** |
| **Edit existing** | Risky round-trip | Good (unpack-pack) | **Good (unpack-pack)** |
| **Validation** | None | formula_check.py | **Both check + audit** |
| **File size** | 391 KB | 20 KB | **33 KB** |
| **Reliability** | Good | Poor (broken sheets) | **Good (all sheets work)** |
| **Financial colors** | Partial | XSD standard | **Both + conditional fills** |

## Quick Start | بداية سريعة

```bash
# Clone into your agent's skills directory
git clone https://github.com/malgaroshy-maker/algaroshy-xslx.git ~/.agents/skills/algaroshy-xlsx

# Or copy directly
cp -r algaroshy-xlsx ~/.agents/skills/
```

Restart your agent to discover the skill. Then ask:

> "Analyze sales.xlsx — give me monthly revenue trends with formulas and charts"

## Skill Architecture

```
algaroshy-xlsx/
├── SKILL.md                    # Core skill instructions
├── scripts/
│   ├── formula_check.py        # Zero-error validation (from MiniMax, MIT)
│   ├── formula_audit.py        # Formula vs hardcode ratio report
│   ├── style_apply.py          # Financial colors + conditional fills
│   ├── xlsx_unpack.py          # OOXML unpack for zero-loss editing
│   └── xlsx_pack.py            # OOXML repack after XML editing
├── references/                 # Detailed guides (extend as needed)
├── demo/
│   ├── example_data.xlsx       # Sample manufacturing dataset
│   └── build_analysis.py       # Example: full 8-sheet analysis workbook
└── LICENSE
```

## Task Routing

| Task | Method | Key Rule |
|------|--------|----------|
| **READ** | `pandas` | Discovery first, then analysis |
| **CREATE** | `openpyxl` + formulas | Every derived value = Excel formula |
| **ADD CHART** | `openpyxl.chart` | Native Excel chart objects |
| **EDIT** | XML unpack→edit→pack | Preserve formatting, touch only target cells |
| **VALIDATE** | `formula_check.py` | Zero formula errors before delivery |
| **AUDIT** | `formula_audit.py` | Formula ratio per sheet |

## Formula-First Philosophy

Every calculated cell **must** be an Excel formula, never a hardcoded number.

```python
# WRONG — computing in Python
total = df['Sales'].sum()
ws['B10'] = total                    # Hardcoded — won't update

# CORRECT — Excel formula
ws['B10'] = '=SUM(B2:B9)'          # Live formula — recalculates
ws['C5'] = '=(C4-C2)/C2'           # Growth rate as formula
ws['D20'] = '=AVERAGE(D2:D19)'     # Average as formula
```

**Exception:** raw input data cells can be hardcoded. Only **derived values** (totals, ratios, percentages, growth) must be formulas.

## Financial Color Standard

| Cell Role | Font Color | When |
|-----------|-----------|------|
| Hard-coded input | Blue `0000FF` | Raw data cells |
| Formula result | Black `000000` | All `=SUM()`, `=AVERAGE()`, etc. |
| Cross-sheet ref | Green `00B050` | Links to other sheets |

## Native Charts (not images)

```python
from openpyxl.chart import BarChart, PieChart, Reference

chart = BarChart()
chart.title = 'Monthly Revenue'
data = Reference(ws, min_col=2, min_row=1, max_row=13)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, 'F3')  # Real Excel chart, not a PNG
```

## Demo: Nexa Manufacturing Analysis

The demo rebuilds a comprehensive 8-sheet analysis workbook from a manufacturing dataset:

| Sheet | Contents |
|-------|----------|
| Executive Summary | KPIs, top 5 actions |
| Revenue Trends | Monthly bar chart, MoM growth, formulas |
| Revenue Region-Product | Pie chart by region, bar chart by product |
| Customer Analysis | Concentration risk, payment delay flags |
| Inventory Analysis | ABC classification, stockout risk scoring |
| Employee Performance | OT efficiency formula, conditional fills |
| Support Tickets | Root cause Pareto, severity breakdown |
| Budget Variance | Dept variance chart, live `=C{n}-B{n}` formulas |

Run it:

```bash
pip install pandas openpyxl
python demo/build_analysis.py
```

Opens `algaroshy-xlsx-analysis.xlsx` — 209 formulas, 4 native charts, 8 sheets.

---

## النسخة العربية | Arabic Version

### ما هي الجروشي إكسلسكس؟

مهارة إكسل مدمجة تجمع أفضل ما في مكتبتين:
- **Anthropic xlsx**: موثوقية openpyxl وتحليل pandas
- **MiniMax minimax-xlsx**: فلسفة الصيغ الحية، ألوان مالية قياسية، أدوات التحقق

### المميزات الرئيسية

- **كل قيمة مشتقة = صيغة إكسل حية** — ليست أرقاماً صلبة. عند تغيير المدخلات، تعيد النتائج الحساب تلقائياً
- **رسوم بيانية أصلية** — ليست صور PNG، بل رسوم Excel حقيقية قابلة للتعديل
- **تدقيق مالي مدمج** — ألوان قياسية (أزرق = مدخلات، أسود = صيغ، أخضر = مرجع خارجي)
- **التحقق من الصيغ** — `formula_check.py` يضمن صفر أخطاء قبل التسليم
- **تقرير نسبة الصيغ** — `formula_audit.py` يظهر نسبة الصيغ مقابل الأرقام الصلبة لكل ورقة

### التثبيت السريع

انسخ هذا السطر في وكيل OpenCode:

```
ثبت مهارة الإكسل هذه في opencode: https://github.com/malgaroshy-maker/algaroshy-xslx.git
```

أو يدوياً:

```bash
git clone https://github.com/malgaroshy-maker/algaroshy-xslx.git ~/.agents/skills/algaroshy-xlsx
```

### مثال: تحليل شركة نكسة للتصنيع

يقوم العرض التوضيحي ببناء مصنف تحليلي من 8 أوراق:

| الورقة | المحتوى |
|--------|---------|
| الملخص التنفيذي | مؤشرات الأداء الرئيسية، أهم 5 إجراءات |
| اتجاهات الإيرادات | رسم بياني شهري، نمو شهري، صيغ حية |
| الإيرادات حسب المنطقة والمنتج | رسم دائري للمناطق، رسم بياني للمنتجات |
| تحليل العملاء | مخاطر التركيز، تأخير الدفع |
| تحليل المخزون | تصنيف ABC، مخاطر النفاد |
| أداء الموظفين | كفاءة العمل الإضافي، تلوين شرطي |
| تذاكر الدعم | تحليل باريتو للأسباب الجذرية |
| انحراف الميزانية | رسم بياني للانحراف، صيغ `=C{n}-B{n}` حية |

```bash
pip install pandas openpyxl
python demo/build_analysis.py
```

## Credits

Built by merging the best of:

- [Anthropic xlsx skill](https://github.com/anthropics/skills/tree/main/skills/xlsx) (Apache 2.0) — openpyxl patterns, pandas analysis
- [MiniMax minimax-xlsx](https://github.com/MiniMax-AI/skills/tree/main/skills/minimax-xlsx) (MIT) — formula-first philosophy, financial colors, `formula_check.py`

## License

MIT — see [LICENSE](LICENSE) for details. Third-party code (formula_check.py, xlsx_unpack.py, xlsx_pack.py) retains original MIT headers.
