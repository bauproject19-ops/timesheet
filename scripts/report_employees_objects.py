"""Ежедневный (19:00 МСК) отчёт: лист «По сотрудникам» (объекты, часы
сегодня/за месяц) и лист «По объектам» (часы всего, кто работал). Собирается
как .xlsx, заливается на Bitrix24 Диск, в чат уходит ссылка на файл."""
from datetime import datetime, timedelta, timezone

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from bitrix_common import (
    USER_NAMES, fmt_num, get_all_items, record_date, send_message,
    upload_to_disk,
)

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
BOLD = Font(bold=True)


def main():
    now_msk = datetime.now(timezone.utc) + timedelta(hours=3)
    today = now_msk.date()
    today_str = today.strftime("%d.%m.%Y")

    items = [it for it in get_all_items() if it.get("ufCrm37Object") != "Отпуск"]
    for it in items:
        it["_date"] = record_date(it.get("title", ""))

    wb = openpyxl.Workbook()

    # --- Лист 1: По сотрудникам ---
    ws1 = wb.active
    ws1.title = "По сотрудникам"
    ws1.cell(row=1, column=1, value=f"Отчёт по сотрудникам — {today_str}").font = Font(bold=True, size=13)
    headers1 = ["Сотрудник", "Объект", "Часы сегодня", "Часы за месяц"]
    for col, h in enumerate(headers1, 1):
        c = ws1.cell(row=2, column=col, value=h)
        c.font = BOLD
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

    by_emp: dict = {}
    for it in items:
        by_emp.setdefault(it["assignedById"], []).append(it)

    row = 3
    grand_today = 0.0
    grand_month = 0.0
    for emp in sorted(by_emp, key=lambda e: USER_NAMES.get(e, str(e))):
        recs = by_emp[emp]
        name = USER_NAMES.get(emp, f"ID{emp}")
        by_obj_today: dict = {}
        month_total = 0.0
        today_total = 0.0
        for it in recs:
            d = it["_date"]
            if not d:
                continue
            h = float(it.get("ufCrm37Hours") or 0)
            if d == today:
                obj = it.get("ufCrm37Object") or "—"
                by_obj_today[obj] = by_obj_today.get(obj, 0.0) + h
                today_total += h
            if d.year == today.year and d.month == today.month:
                month_total += h

        if by_obj_today:
            for obj, h in by_obj_today.items():
                ws1.cell(row=row, column=1, value=name)
                ws1.cell(row=row, column=2, value=obj)
                ws1.cell(row=row, column=3, value=fmt_num(h))
                ws1.cell(row=row, column=4, value="")
                row += 1
        else:
            ws1.cell(row=row, column=1, value=name)
            ws1.cell(row=row, column=2, value="нет записей за сегодня")
            ws1.cell(row=row, column=3, value=0)
            ws1.cell(row=row, column=4, value="")
            row += 1

        ws1.cell(row=row, column=1, value=f"Всего — {name}")
        ws1.cell(row=row, column=2, value="")
        ws1.cell(row=row, column=3, value=fmt_num(today_total))
        ws1.cell(row=row, column=4, value=fmt_num(month_total))
        for col in range(1, 5):
            ws1.cell(row=row, column=col).font = BOLD
            ws1.cell(row=row, column=col).fill = TOTAL_FILL
        row += 1
        grand_today += today_total
        grand_month += month_total

    ws1.cell(row=row, column=1, value="ИТОГО").font = BOLD
    ws1.cell(row=row, column=3, value=fmt_num(grand_today)).font = BOLD
    ws1.cell(row=row, column=4, value=fmt_num(grand_month)).font = BOLD

    for r in range(2, row + 1):
        for col in range(1, 5):
            ws1.cell(row=r, column=col).border = BORDER

    ws1.column_dimensions["A"].width = 24
    ws1.column_dimensions["B"].width = 28
    ws1.column_dimensions["C"].width = 14
    ws1.column_dimensions["D"].width = 14

    # --- Лист 2: По объектам ---
    ws2 = wb.create_sheet("По объектам")
    ws2.cell(row=1, column=1, value=f"Отчёт по объектам — всего (на {today_str})").font = Font(bold=True, size=13)
    headers2 = ["Объект", "Часы всего", "Сотрудники"]
    for col, h in enumerate(headers2, 1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font = BOLD
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

    by_obj: dict = {}
    for it in items:
        obj = it.get("ufCrm37Object")
        if not obj:
            continue
        h = float(it.get("ufCrm37Hours") or 0)
        d = by_obj.setdefault(obj, {"hours": 0.0, "emps": set()})
        d["hours"] += h
        d["emps"].add(USER_NAMES.get(it["assignedById"], f"ID{it['assignedById']}"))

    row = 3
    grand = 0.0
    for obj, d in sorted(by_obj.items(), key=lambda kv: -kv[1]["hours"]):
        ws2.cell(row=row, column=1, value=obj)
        ws2.cell(row=row, column=2, value=fmt_num(d["hours"]))
        ws2.cell(row=row, column=3, value=", ".join(sorted(d["emps"])))
        for col in range(1, 4):
            ws2.cell(row=row, column=col).border = BORDER
        grand += d["hours"]
        row += 1

    ws2.cell(row=row, column=1, value="ИТОГО").font = BOLD
    ws2.cell(row=row, column=2, value=fmt_num(grand)).font = BOLD
    for col in range(1, 4):
        ws2.cell(row=row, column=col).border = BORDER
        ws2.cell(row=row, column=col).fill = TOTAL_FILL

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 40

    filename = f"Отчет_по_сотрудникам_и_объектам_{today_str}.xlsx"
    wb.save(filename)
    with open(filename, "rb") as f:
        content = f.read()
    url = upload_to_disk(filename, content)
    send_message("chat66269", f"[B]📊 Отчёт по сотрудникам и объектам — {today_str}[/B]\n{url}")
    print(f"OK: {url}")


if __name__ == "__main__":
    main()
