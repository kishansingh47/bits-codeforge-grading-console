# Test Files

Sample `.xlsx` files for exercising the grading console, each targeting a specific scenario. Upload any of these at `BITS_Digital_CodeForge_Challenge.html` and try the corresponding checks.

| File | Scenario | What to check |
|---|---|---|
| `sample_marks.xlsx` | Clean, everyday upload — 2 courses (A: 8 students, B: 3 students) | Course dropdown has no duplicates; stats/histogram/grade counts match the data |
| `boundary_marks.xlsx` | One student at every single grade-band edge (0, 19, 20, 29, 30, 39, ... 80, 100) | Every grade (A–E) gets exactly 2 students, confirming no gaps or overlaps at the boundaries |
| `single_student.xlsx` | Only one row in the course | Max/Min/Avg/Median all equal that one mark; nothing shows `NaN` or crashes |
| `zero_variance.xlsx` | 12 students, all with the identical mark (60) | Standard deviation is 0 — histogram/bell-curve and stats still render without errors |
| `large_class.xlsx` | 150 students, random marks 0–100 | Histogram, stats and the results table stay responsive with a larger dataset |
| `multi_course.xlsx` | 5 courses in one file, uneven sizes (1–10 students each) | Dropdown lists all 5 courses once each; switching between them shows the right subset |
| `missing_columns.xlsx` | Headers renamed (`Student ID`, `Marks` instead of `BITS ID`, `Total Marks`) | Upload is rejected with a clear "missing required column(s)" warning; course list stays empty |
| `empty_data.xlsx` | Header row only, zero data rows | Upload shows a "no data rows" warning instead of silently doing nothing |
| `non_numeric_marks.xlsx` | Mix of valid marks and invalid ones (`"Absent"`, `"N/A"`, blank) | The 2 valid rows are kept; the 3 invalid rows are dropped and counted in the warning banner |
| `dirty_marks.xlsx` | Fractional mark, out-of-range marks (-5, 105), missing BITS ID, duplicate BITS ID | Fractional mark is rounded; out-of-range/missing-ID rows are dropped; duplicate is flagged — all reported in the warning banner |

Regenerate these at any time with the script below (requires `openpyxl`):

```bash
python3 -c "
import openpyxl, random
def save(name, header, rows):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(header)
    for r in rows: ws.append(r)
    wb.save(f'testdata/{name}.xlsx')
"
```

(See the full generation script in the project's commit history for the exact data used above.)
