# -*- coding: utf-8 -*-
"""يولّد ملف students.js من ملف Excel لوائح التلاميذ (مسار: ListEleve_*.xlsx).
الاستعمال:  python build_students.py ListEleve_20260914.xlsx
لا يحفظ إلا: النسب، الاسم، القسم. (لا رمز مسار، لا تاريخ ازدياد، لا مكان ازدياد)
"""
import sys, json, re, openpyxl

src = sys.argv[1] if len(sys.argv) > 1 else "ListEleve_20260914.xlsx"
clean = lambda v: re.sub(r"\s+", " ", str(v or "")).strip()

wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
rows, classes = [], {}
for ws in wb:
    data = list(ws.iter_rows(values_only=True))
    if len(data) < 11:
        continue
    level = clean(data[6][2]).replace(" عام", "")
    code = clean(data[7][2]) or ws.title
    classes[code] = level
    for r in data[10:]:
        nom, pre = clean(r[2]), clean(r[3])
        if nom or pre:
            rows.append([nom, pre, code])

out = "window.SALAM_CLASSES=" + json.dumps(classes, ensure_ascii=False) + ";\n"
out += "window.SALAM_ELEVES=[\n" + ",\n".join(json.dumps(r, ensure_ascii=False) for r in rows) + "\n];\n"
open("students.js", "w", encoding="utf-8").write(out)
print(len(rows), "تلميذ،", len(classes), "قسم ->", "students.js")
