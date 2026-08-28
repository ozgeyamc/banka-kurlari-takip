from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FILL = PatternFill("solid", fgColor="D9EAF7")
TITLE_FONT = Font(bold=True, size=14, color="1F1F1F")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)
OK_FILL = PatternFill("solid", fgColor="E2F0D9")
CONTROL_FILL = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL = PatternFill("solid", fgColor="F4CCCC")


def _parse_dt(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _to_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def read_history(path: str | Path) -> list[dict]:
    path = Path(path)
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _style_header(ws, row: int, start_col: int, end_col: int) -> None:
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _apply_table(ws, start_row: int, end_row: int, end_col: int, name: str) -> None:
    if end_row <= start_row:
        return

    from openpyxl.utils import get_column_letter

    ref = f"A{start_row}:{get_column_letter(end_col)}{end_row}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def _set_widths(ws, widths: dict[str, float]) -> None:
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def _status_fill(status: str):
    if status == "ERROR":
        return ERROR_FILL
    if status == "KONTROL":
        return CONTROL_FILL
    return OK_FILL


def _latest_run_rows(history: list[dict]) -> tuple[str | None, list[dict]]:
    valid = [row for row in history if _parse_dt(row.get("run_at"))]
    if not valid:
        return None, []

    latest = max(valid, key=lambda row: _parse_dt(row.get("run_at")))
    latest_run_at = latest["run_at"]
    return latest_run_at, [row for row in history if row.get("run_at") == latest_run_at]


def _provider_map(rows: Iterable[dict]) -> dict[str, dict[str, dict]]:
    result: dict[str, dict[str, dict]] = {}
    for row in rows:
        provider = row.get("provider", "").strip()
        code = row.get("code", "").strip()
        if not provider or not code:
            continue
        result.setdefault(provider, {})[code] = row
    return result


def _build_summary_sheet(wb: Workbook, latest_run_at: str | None, latest_rows: list[dict]) -> None:
    ws = wb.create_sheet("OZET")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    ws["A1"] = "Döviz.com Kur Takip Özeti"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    run_dt = _parse_dt(latest_run_at)
    providers = {row.get("provider") for row in latest_rows if row.get("provider")}
    error_count = sum(row.get("status") == "ERROR" for row in latest_rows)
    control_count = sum(row.get("status") == "KONTROL" for row in latest_rows)

    labels = [
        ("A3", "Son Çekim Tarihi"),
        ("A4", "Son Çekim Saati"),
        ("A5", "Toplam Sağlayıcı"),
        ("A6", "Son Çekim Toplam Kayıt"),
        ("A7", "ERROR"),
        ("A8", "KONTROL"),
    ]
    for coord, label in labels:
        ws[coord] = label
        ws[coord].font = Font(bold=True)
        ws[coord].fill = PatternFill("solid", fgColor="EAF2F8")
        ws[coord].border = THIN_BORDER

    ws["B3"] = run_dt.date() if run_dt else ""
    ws["B4"] = run_dt.time().replace(tzinfo=None) if run_dt else ""
    ws["B5"] = len(providers)
    ws["B6"] = len(latest_rows)
    ws["B7"] = error_count
    ws["B8"] = control_count
    ws["B3"].number_format = "dd.mm.yyyy"
    ws["B4"].number_format = "hh:mm:ss"
    for row in range(3, 9):
        ws[f"B{row}"].border = THIN_BORDER

    headers = ["Ürün", "Sağlayıcı Sayısı", "En Düşük Makas %", "Sağlayıcı", "Alış", "Satış"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=11, column=col, value=header)
    _style_header(ws, 11, 1, len(headers))

    code_names = {"USD": "DOLAR", "EUR": "EURO", "XAU": "GRAM ALTIN"}
    row_no = 12
    for code in ("USD", "EUR", "XAU"):
        product_rows = [row for row in latest_rows if row.get("code") == code]
        valid_rows = []
        for row in product_rows:
            buy = _to_float(row.get("buy"))
            sell = _to_float(row.get("sell"))
            pct = _to_float(row.get("spread_pct"))
            if row.get("status") != "ERROR" and buy and sell and pct is not None:
                valid_rows.append((pct, row, buy, sell))

        best = min(valid_rows, key=lambda item: item[0]) if valid_rows else None
        ws.cell(row=row_no, column=1, value=code_names[code])
        ws.cell(row=row_no, column=2, value=len(product_rows))
        if best:
            pct, row, buy, sell = best
            ws.cell(row=row_no, column=3, value=pct / 100.0)
            ws.cell(row=row_no, column=4, value=row.get("provider"))
            ws.cell(row=row_no, column=5, value=buy)
            ws.cell(row=row_no, column=6, value=sell)
            ws.cell(row=row_no, column=3).number_format = "0.00%"
            ws.cell(row=row_no, column=5).number_format = "#,##0.0000"
            ws.cell(row=row_no, column=6).number_format = "#,##0.0000"
        for col in range(1, 7):
            ws.cell(row=row_no, column=col).border = THIN_BORDER
        row_no += 1

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Ürün Bazında En Düşük Makas %"
    chart.y_axis.title = "Makas %"
    chart.x_axis.title = "Ürün"
    data = Reference(ws, min_col=3, min_row=11, max_row=14)
    cats = Reference(ws, min_col=1, min_row=12, max_row=14)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 7
    chart.width = 12
    ws.add_chart(chart, "H3")

    _set_widths(ws, {"A": 22, "B": 18, "C": 18, "D": 24, "E": 16, "F": 16})


def _build_current_sheet(wb: Workbook, latest_run_at: str | None, latest_rows: list[dict]) -> None:
    ws = wb.create_sheet("GUNCEL_KURLAR")
    ws.freeze_panes = "D2"
    ws.sheet_view.showGridLines = False

    headers = [
        "Tarih",
        "Saat",
        "Kurum / Sağlayıcı",
        "Dolar Alış",
        "Dolar Satış",
        "Dolar Makas",
        "Dolar Makas %",
        "Euro Alış",
        "Euro Satış",
        "Euro Makas",
        "Euro Makas %",
        "Gram Altın Alış",
        "Gram Altın Satış",
        "Gram Altın Makas",
        "Gram Altın Makas %",
    ]
    ws.append(headers)
    _style_header(ws, 1, 1, len(headers))

    provider_map = _provider_map(latest_rows)
    run_dt = _parse_dt(latest_run_at)

    for excel_row, provider in enumerate(sorted(provider_map, key=str.casefold), start=2):
        row_map = provider_map[provider]
        ws.cell(excel_row, 1, run_dt.date() if run_dt else "")
        ws.cell(excel_row, 2, run_dt.time().replace(tzinfo=None) if run_dt else "")
        ws.cell(excel_row, 3, provider)

        layout = {
            "USD": (4, 5, 6, 7),
            "EUR": (8, 9, 10, 11),
            "XAU": (12, 13, 14, 15),
        }

        for code, (buy_col, sell_col, spread_col, pct_col) in layout.items():
            item = row_map.get(code)
            if not item:
                continue

            buy = _to_float(item.get("buy"))
            sell = _to_float(item.get("sell"))
            if buy is not None:
                ws.cell(excel_row, buy_col, buy)
            if sell is not None:
                ws.cell(excel_row, sell_col, sell)

            buy_letter = ws.cell(1, buy_col).column_letter
            sell_letter = ws.cell(1, sell_col).column_letter
            spread_letter = ws.cell(1, spread_col).column_letter

            ws.cell(
                excel_row,
                spread_col,
                f'=IF(OR({buy_letter}{excel_row}="",{sell_letter}{excel_row}=""),"",{sell_letter}{excel_row}-{buy_letter}{excel_row})',
            )
            ws.cell(
                excel_row,
                pct_col,
                f'=IFERROR({spread_letter}{excel_row}/{buy_letter}{excel_row},"")',
            )

        ws.cell(excel_row, 1).number_format = "dd.mm.yyyy"
        ws.cell(excel_row, 2).number_format = "hh:mm:ss"

        for col in (4, 5, 6, 8, 9, 10, 12, 13, 14):
            ws.cell(excel_row, col).number_format = "#,##0.0000"
        for col in (7, 11, 15):
            ws.cell(excel_row, col).number_format = "0.00%"

    if ws.max_row >= 2:
        _apply_table(ws, 1, ws.max_row, len(headers), "GuncelKurlarTable")

    _set_widths(
        ws,
        {
            "A": 13,
            "B": 12,
            "C": 26,
            "D": 15,
            "E": 15,
            "F": 15,
            "G": 14,
            "H": 15,
            "I": 15,
            "J": 15,
            "K": 14,
            "L": 18,
            "M": 18,
            "N": 18,
            "O": 17,
        },
    )


def _build_history_sheet(wb: Workbook, history: list[dict]) -> None:
    ws = wb.create_sheet("GECMIS")
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    headers = [
        "Tarih",
        "Saat",
        "Kurum / Sağlayıcı",
        "Ürün",
        "Alış",
        "Satış",
        "Makas",
        "Makas %",
        "Durum",
        "Kaynak",
        "Sitedeki Makas",
        "Sitedeki Makas %",
        "Not",
        "Gerçek Ürün Çekim Saati",
    ]
    ws.append(headers)
    _style_header(ws, 1, 1, len(headers))

    sorted_history = sorted(
        history,
        key=lambda row: (_parse_dt(row.get("run_at")) or datetime.min, row.get("provider", ""), row.get("code", "")),
    )

    for excel_row, item in enumerate(sorted_history, start=2):
        run_dt = _parse_dt(item.get("run_at"))
        scraped_dt = _parse_dt(item.get("scraped_at"))
        buy = _to_float(item.get("buy"))
        sell = _to_float(item.get("sell"))
        site_spread = _to_float(item.get("site_spread"))
        site_spread_pct = _to_float(item.get("site_spread_pct"))

        ws.cell(excel_row, 1, run_dt.date() if run_dt else "")
        ws.cell(excel_row, 2, run_dt.time().replace(tzinfo=None) if run_dt else "")
        ws.cell(excel_row, 3, item.get("provider", ""))
        ws.cell(excel_row, 4, item.get("product", ""))
        ws.cell(excel_row, 5, buy if buy is not None else "")
        ws.cell(excel_row, 6, sell if sell is not None else "")
        ws.cell(excel_row, 7, f'=IF(OR(E{excel_row}="",F{excel_row}=""),"",F{excel_row}-E{excel_row})')
        ws.cell(excel_row, 8, f'=IFERROR(G{excel_row}/E{excel_row},"")')
        ws.cell(excel_row, 9, item.get("status", ""))
        ws.cell(excel_row, 10, item.get("source_url", ""))
        ws.cell(excel_row, 11, site_spread if site_spread is not None else "")
        ws.cell(excel_row, 12, site_spread_pct / 100.0 if site_spread_pct is not None else "")
        ws.cell(excel_row, 13, item.get("note", ""))
        ws.cell(excel_row, 14, scraped_dt.time().replace(tzinfo=None) if scraped_dt else "")

        ws.cell(excel_row, 1).number_format = "dd.mm.yyyy"
        ws.cell(excel_row, 2).number_format = "hh:mm:ss"
        ws.cell(excel_row, 14).number_format = "hh:mm:ss"
        for col in (5, 6, 7, 11):
            ws.cell(excel_row, col).number_format = "#,##0.0000"
        for col in (8, 12):
            ws.cell(excel_row, col).number_format = "0.00%"

        source_cell = ws.cell(excel_row, 10)
        if source_cell.value:
            source_cell.hyperlink = source_cell.value
            source_cell.style = "Hyperlink"

        status_cell = ws.cell(excel_row, 9)
        status_cell.fill = _status_fill(str(status_cell.value))
        status_cell.font = Font(bold=True)

    if ws.max_row >= 2:
        _apply_table(ws, 1, ws.max_row, len(headers), "GecmisTable")

    _set_widths(
        ws,
        {
            "A": 13,
            "B": 12,
            "C": 26,
            "D": 16,
            "E": 15,
            "F": 15,
            "G": 15,
            "H": 14,
            "I": 12,
            "J": 45,
            "K": 17,
            "L": 18,
            "M": 60,
            "N": 22,
        },
    )
    ws.column_dimensions["M"].width = 60
    for row in ws.iter_rows(min_row=2, min_col=13, max_col=13):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")


def _build_control_sheet(wb: Workbook, latest_run_at: str | None, latest_rows: list[dict]) -> None:
    ws = wb.create_sheet("KONTROL")
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    headers = [
        "Tarih",
        "Saat",
        "Ürün",
        "Kurum / Sağlayıcı",
        "Durum",
        "Not",
        "Alış",
        "Satış",
        "Hesaplanan Makas %",
        "Sitedeki Makas %",
        "Kaynak",
    ]
    ws.append(headers)
    _style_header(ws, 1, 1, len(headers))

    run_dt = _parse_dt(latest_run_at)
    control_rows = [row for row in latest_rows if row.get("status") != "OK"]

    for excel_row, item in enumerate(control_rows, start=2):
        pct = _to_float(item.get("spread_pct"))
        site_pct = _to_float(item.get("site_spread_pct"))

        ws.cell(excel_row, 1, run_dt.date() if run_dt else "")
        ws.cell(excel_row, 2, run_dt.time().replace(tzinfo=None) if run_dt else "")
        ws.cell(excel_row, 3, item.get("product", ""))
        ws.cell(excel_row, 4, item.get("provider", ""))
        ws.cell(excel_row, 5, item.get("status", ""))
        ws.cell(excel_row, 6, item.get("note", ""))
        ws.cell(excel_row, 7, _to_float(item.get("buy")) or "")
        ws.cell(excel_row, 8, _to_float(item.get("sell")) or "")
        ws.cell(excel_row, 9, pct / 100.0 if pct is not None else "")
        ws.cell(excel_row, 10, site_pct / 100.0 if site_pct is not None else "")
        ws.cell(excel_row, 11, item.get("source_url", ""))

        ws.cell(excel_row, 1).number_format = "dd.mm.yyyy"
        ws.cell(excel_row, 2).number_format = "hh:mm:ss"
        for col in (7, 8):
            ws.cell(excel_row, col).number_format = "#,##0.0000"
        for col in (9, 10):
            ws.cell(excel_row, col).number_format = "0.00%"

        status_cell = ws.cell(excel_row, 5)
        status_cell.fill = _status_fill(str(status_cell.value))
        status_cell.font = Font(bold=True)

        source_cell = ws.cell(excel_row, 11)
        if source_cell.value:
            source_cell.hyperlink = source_cell.value
            source_cell.style = "Hyperlink"

    if ws.max_row == 1:
        ws.cell(2, 1, "Son çekimde ERROR veya KONTROL kaydı yok.")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        ws["A2"].fill = OK_FILL
        ws["A2"].font = Font(bold=True)
        ws["A2"].alignment = Alignment(horizontal="center")
    else:
        _apply_table(ws, 1, ws.max_row, len(headers), "KontrolTable")

    _set_widths(
        ws,
        {
            "A": 13,
            "B": 12,
            "C": 16,
            "D": 26,
            "E": 12,
            "F": 60,
            "G": 15,
            "H": 15,
            "I": 20,
            "J": 20,
            "K": 45,
        },
    )
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")


def build_excel(history_path: str | Path, output_path: str | Path) -> None:
    history = read_history(history_path)
    if not history:
        raise RuntimeError("Excel oluşturmak için geçmiş veri bulunamadı.")

    latest_run_at, latest_rows = _latest_run_rows(history)

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    _build_summary_sheet(wb, latest_run_at, latest_rows)
    _build_current_sheet(wb, latest_run_at, latest_rows)
    _build_history_sheet(wb, history)
    _build_control_sheet(wb, latest_run_at, latest_rows)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
