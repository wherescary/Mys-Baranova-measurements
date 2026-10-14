#!/usr/bin/env python3
"""
Разбор ежемесячных отчётов магнитной станции «Мыс Баранова» (.docx).

Каждый отчёт содержит две таблицы:
  Табл.1 — базисные значения вариометра (D, H, Z, T)
  Табл.2 — абсолютные значения МПЗ      (D, H, Z, I, T)
плюс строки «среднее» и «СКО».

Запуск:  python3 parse.py reports/  site/
Результат: site/data.json, site/data.csv, site/check_report.txt
"""
import csv
import json
import math
import re
import statistics
import sys
from datetime import datetime
from pathlib import Path

import docx  # pip install python-docx

MONTH_NAMES = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
               "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
MONTHS = {
    "ЯНВАРЬ": 1, "ФЕВРАЛЬ": 2, "МАРТ": 3, "АПРЕЛЬ": 4, "МАЙ": 5, "ИЮНЬ": 6,
    "ИЮЛЬ": 7, "АВГУСТ": 8, "СЕНТЯБРЬ": 9, "ОКТЯБРЬ": 10, "НОЯБРЬ": 11, "ДЕКАБРЬ": 12,
}
BASIS_COLS = ["D", "H", "Z", "T"]
ABS_COLS = ["D", "H", "Z", "I", "T"]


def num(s):
    """Число из ячейки; None, если ячейка пустая или там не число (например, повторена шапка)."""
    s = s.strip().replace(",", ".").replace("\u2212", "-").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


def table_rows(table):
    """Строки таблицы как списки текстов (объединённые ячейки не дублируются)."""
    out = []
    for r in table.rows:
        cells, prev = [], None
        for c in r.cells:
            if c._tc is prev:
                continue
            prev = c._tc
            cells.append(c.text.strip())
        out.append(cells)
    return out


def parse_table(table, cols):
    """Возвращает (измерения, среднее, СКО) из одной таблицы отчёта."""
    rows, mean, std = [], None, None
    for cells in table_rows(table)[1:]:  # первая строка — заголовок
        if not any(cells):
            continue
        label = (cells[0] + " " + cells[1]).lower()
        values = [num(x) for x in cells[2:2 + len(cols)]]
        if re.search(r"ср|сред", label) and not re.search(r"\d\d\.\d\d", cells[0]):
            mean = dict(zip(cols, values))
        elif re.search(r"откл", label):
            std = dict(zip(cols, values))
        elif re.match(r"\d{1,2}\.\d{1,2}\.\d{2,4}", cells[0]):
            rows.append({"date": cells[0], "time": cells[1], **dict(zip(cols, values))})
    return rows, mean, std


def parse_report(path):
    d = docx.Document(path)
    text = "\n".join(p.text for p in d.paragraphs)
    m = re.search(r"ЗА\s+([А-ЯЁ]+)\s+(\d{4})", text.upper())
    if not m:
        raise ValueError("не найден месяц отчёта («ЗА <МЕСЯЦ> <ГОД> Г.»)")
    month, year = MONTHS[m.group(1)], int(m.group(2))
    inst = re.search(r"Базисные значения\s+(.+)", text)
    instrument = inst.group(1).replace("МВС", "").strip() if inst else ""
    days = re.search(r"суточных магнитограмм:\s*(\d+)", text)
    if len(d.tables) < 2:
        raise ValueError(f"ожидалось 2 таблицы, найдено {len(d.tables)}")
    basis, b_mean, b_std = parse_table(d.tables[0], BASIS_COLS)
    absol, a_mean, a_std = parse_table(d.tables[1], ABS_COLS)
    return {
        "file": Path(path).name, "year": year, "month": month,
        "instrument": instrument, "days": int(days.group(1)) if days else None,
        "basis": basis, "abs": absol,
        "basis_mean": b_mean, "basis_std": b_std, "abs_mean": a_mean, "abs_std": a_std,
    }


def to_datetime(date_s, time_s, year, month, issues, where):
    dd, mm, yy = date_s.split(".")
    yy = int(yy) + 2000 if len(yy) == 2 else int(yy)
    dd, mm = int(dd), int(mm)
    if (yy, mm) != (year, month):
        if mm == month:
            issues.append(f"{where}: дата {date_s} не совпадает с годом отчёта, исправлено на {year}")
            yy = year
        else:
            issues.append(f"{where}: дата {date_s} вне месяца отчёта {month:02d}.{year}")
    h, mi, s = (int(x) for x in time_s.split(":"))
    return datetime(yy, mm, dd, h, mi, s)


def main(src, dst):
    src, dst = Path(src), Path(dst)
    dst.mkdir(parents=True, exist_ok=True)
    issues, records, months, seen = [], {}, {}, {}

    for path in sorted(src.glob("*.docx")):
        if path.name.startswith("~$"):
            continue
        try:
            rep = parse_report(path)
        except Exception as e:
            issues.append(f"{path.name}: НЕ ПРОЧИТАН — {e}")
            continue
        key = (rep["year"], rep["month"])
        tag = f"Отчёт за {MONTH_NAMES[rep['month'] - 1]} {rep['year']}"
        if key in seen:
            issues.append(f"{path.name}: дубликат отчёта за {rep['month']:02d}.{rep['year']} (уже есть {seen[key]}), пропущен")
            continue
        seen[key] = path.name
        if len(rep["basis"]) != len(rep["abs"]):
            issues.append(f"{tag}: разное число строк в таблицах ({len(rep['basis'])} и {len(rep['abs'])})")

        for b, a in zip(rep["basis"], rep["abs"]):
            if (b["date"], b["time"]) != (a["date"], a["time"]):
                issues.append(f"{tag}: даты в таблицах не совпадают ({b['date']} {b['time']} / {a['date']} {a['time']})")
            n_before = len(issues)
            dt = to_datetime(a["date"], a["time"], rep["year"], rep["month"], issues, tag)
            flags = [m.split(": ", 1)[1] for m in issues[n_before:]]
            calc_flags = []
            for name, row, cols in (("табл. 1", b, BASIS_COLS), ("табл. 2", a, ABS_COLS)):
                missing = [c for c in cols if row[c] is None]
                if missing:
                    calc_flags.append(f"нет чисел в {name} ({', '.join(missing)}): в ячейках текст вместо значений")
            # проверки согласованности абсолютных значений
            if all(a[c] is not None for c in ("H", "Z", "T", "I")):
                T_calc = math.hypot(a["H"], a["Z"])
                I_calc = math.degrees(math.atan2(a["Z"], a["H"]))
                if abs(T_calc - a["T"]) > 0.05:
                    calc_flags.append(f"T≠√(H²+Z²): {a['T']} vs {T_calc:.2f}")
                if abs(I_calc - a["I"]) > 0.001:
                    calc_flags.append(f"I≠arctg(Z/H): {a['I']} vs {I_calc:.4f}")
            for f in calc_flags:
                issues.append(f"{tag}: {a['date']} {f}")
            flags += calc_flags
            rec = {
                "t": dt.strftime("%Y-%m-%dT%H:%M:%S"),
                **{f"b{k}": b[k] for k in BASIS_COLS},
                **{f"a{k}": a[k] for k in ABS_COLS},
                "instrument": rep["instrument"], "file": path.name,
                "flags": flags,
            }
            if rec["t"] in records:
                issues.append(f"{tag}: повтор измерения {rec['t']}, пропущено")
                continue
            records[rec["t"]] = rec

        # месячные средние/СКО: из отчёта, со сверкой с пересчётом
        stats = {}
        for prefix, rows, cols, mean, std in (
            ("b", rep["basis"], BASIS_COLS, rep["basis_mean"], rep["basis_std"]),
            ("a", rep["abs"], ABS_COLS, rep["abs_mean"], rep["abs_std"]),
        ):
            for c in cols:
                vals = [r[c] for r in rows if r[c] is not None]
                m_calc = statistics.mean(vals) if vals else None
                s_calc = statistics.stdev(vals) if len(vals) > 1 else None
                m_rep = mean.get(c) if mean else None
                s_rep = std.get(c) if std else None
                tol = 0.0002 if c in ("D", "I") else 0.02
                if c == "I":
                    tol = 0.006  # в отчётах среднее и СКО наклонения округлены до 0.01
                name = ("базисное " if prefix == "b" else "абсолютное ") + c
                if m_rep is not None and m_calc is not None and abs(m_rep - m_calc) > tol:
                    issues.append(f"{tag}: среднее {name} в отчёте {m_rep}, по строкам таблицы {m_calc:.4f}")
                if s_rep is not None and s_calc is not None and abs(s_rep - s_calc) > tol:
                    issues.append(f"{tag}: СКО {name} в отчёте {s_rep}, по строкам таблицы {s_calc:.4f}")
                stats[f"{prefix}{c}"] = {
                    "mean": round(m_calc, 4) if m_calc is not None else m_rep,
                    "std": round(s_calc, 4) if s_calc is not None else s_rep,
                }
        months[f"{rep['year']}-{rep['month']:02d}"] = {
            "instrument": rep["instrument"], "days": rep["days"], "n": len(rep["abs"]),
            "file": path.name, "stats": stats,
        }

    data = {
        "station": "Мыс Баранова",
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "records": sorted(records.values(), key=lambda r: r["t"]),
        "months": dict(sorted(months.items())),
        "issues": issues,
    }
    (dst / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    cols = ["t"] + [f"b{k}" for k in BASIS_COLS] + [f"a{k}" for k in ABS_COLS] + ["instrument", "file"]
    with open(dst / "data.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(cols)
        for r in data["records"]:
            w.writerow([r[c] for c in cols])

    report = [f"Отчётов: {len(months)}, измерений: {len(records)}", ""] + (issues or ["Замечаний нет"])
    (dst / "check_report.txt").write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "reports", sys.argv[2] if len(sys.argv) > 2 else "site")
