from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.data_source import NumData, NumVal, StrData, StrRef, StrVal
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# ============================================================
# GENEL AYARLAR
# ============================================================

TARGET_BANKS = [
    "Garanti BBVA",
    "Akbank",
    "Yapıkredi",
    "Ziraat Bankası",
    "İş Bankası",
]

PRODUCT_ORDER = {
    "USD": 0,
    "EUR": 1,
    "XAU": 2,
}

PRODUCT_NAMES = {
    "USD": "DOLAR",
    "EUR": "EURO",
    "XAU": "GRAM ALTIN",
}

MONTH_NAMES = {
    1: "Ocak",
    2: "Şubat",
    3: "Mart",
    4: "Nisan",
    5: "Mayıs",
    6: "Haziran",
    7: "Temmuz",
    8: "Ağustos",
    9: "Eylül",
    10: "Ekim",
    11: "Kasım",
    12: "Aralık",
}


# ============================================================
# RENKLER
# ============================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)

TITLE_FILL = PatternFill("solid", fgColor="D9EAF7")

USD_FILL = PatternFill("solid", fgColor="DDEBF7")
EUR_FILL = PatternFill("solid", fgColor="FCE4D6")
XAU_FILL = PatternFill("solid", fgColor="FFF2CC")

USD_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
EUR_HEADER_FILL = PatternFill("solid", fgColor="ED7D31")
XAU_HEADER_FILL = PatternFill("solid", fgColor="BF9000")

WEEKEND_FILL = PatternFill("solid", fgColor="FFE699")

OK_FILL = PatternFill("solid", fgColor="E2F0D9")
CONTROL_FILL = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL = PatternFill("solid", fgColor="FCE4D6")

BEST_FILL = PatternFill("solid", fgColor="E2F0D9")
BEST_FONT = Font(color="006100", bold=True)

THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)

PRODUCT_LINE_COLORS = {
    "USD": "4472C4",
    "EUR": "ED7D31",
    "XAU": "FFC000",
}

BANK_LINE_COLORS = {
    "Garanti BBVA": "70AD47",
    "Akbank": "C00000",
    "Yapıkredi": "7030A0",
    "Ziraat Bankası": "FFC000",
    "İş Bankası": "4472C4",
}

BANK_CHART_COLORS = {
    "Garanti BBVA": "E2F0D9",
    "Akbank": "FDE2E2",
    "Yapıkredi": "E4D7F5",
    "Ziraat Bankası": "FFF2CC",
    "İş Bankası": "DDEBF7",
}

PROVIDER_COLORS = {
    "Akbank": "ECCACA",
    "Albaraka Türk": "DFF1E4",
    "Alternatif Bank": "DECDEA",
    "Altınkaynak": "F3F0DD",
    "Anadolubank": "CDE5EA",
    "CEPTETEB": "F1DFE9",
    "Denizbank": "D3ECCA",
    "DestekBank": "E0DFF1",
    "Dünya Katılım": "EAD7CD",
    "Emlak Katılım": "DDF3EB",
    "Enpara": "E7CDEA",
    "Fibabanka": "EDF1DF",
    "Garanti BBVA": "CADBEC",
    "Getirfinans": "F1DFE3",
    "Halkbank": "CDEACF",
    "Harem": "E5DDF3",
    "Hayat Finans": "EAE0CD",
    "Hepsipay": "DFF1F1",
    "HSBC": "ECCAE4",
    "ING Bank": "E7F1DF",
    "İş Bankası": "CDD2EA",
    "Kapalıçarşı": "F3E0DD",
    "Kuveyt Türk": "CDEAD9",
    "Merkez Bankası": "ECDFF1",
    "Misyon Bank": "ECECCA",
    "Odacı": "DFECF1",
    "Odeabank": "EACDD9",
    "Papara": "E0F3DD",
    "QNB Finansbank": "D2CDEA",
    "TOM Bank Hadi": "F1E7DF",
    "Türkiye Finans": "CAECE4",
    "Vakıf Katılım": "F1DFF0",
    "Vakıfbank": "E0EACD",
    "Venüs": "DDE5F3",
    "Yapıkredi": "D9C2E9",
    "Ziraat Bankası": "DFF1E3",
    "Ziraat Dinamik": "DBCAEC",
    "Ziraat Katılım": "F1EDDF",
}

DEFAULT_PROVIDER_COLOR = "E8EDF3"
DEFAULT_BANK_CHART_COLOR = "E8EDF3"

STATUS_LABELS = {
    "OK": "DOĞRU",
    "KONTROL": "KONTROL GEREKLİ",
    "ERROR": "HATA",
}


# ============================================================
# GENEL YARDIMCILAR
# ============================================================

def _provider_fill(provider):
    color = PROVIDER_COLORS.get(
        (provider or "").strip(),
        DEFAULT_PROVIDER_COLOR,
    )
    return PatternFill("solid", fgColor=color)


def _product_fill(code):
    code = (code or "").strip().upper()

    if code == "USD":
        return USD_FILL
    if code == "EUR":
        return EUR_FILL
    if code == "XAU":
        return XAU_FILL

    return PatternFill(fill_type=None)


def _parse_dt(value):

    if not value:
        return None

    value = str(value).strip()

    try:
        dt = datetime.fromisoformat(
            value.replace("Z", "+00:00")
        )

        if dt.tzinfo:
            dt = dt.replace(tzinfo=None)

        return dt

    except Exception:
        pass

    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    return None


def _to_float(value):

    if value in (None, ""):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("%", "").replace("\u00a0", "")

    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        pass

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def _get_spread_values(row):

    buy = _to_float(row.get("buy"))
    sell = _to_float(row.get("sell"))
    spread = _to_float(row.get("spread"))
    pct = _to_float(row.get("spread_pct"))

    if spread is None and buy is not None and sell is not None:
        spread = sell - buy

    if pct is None and spread is not None and buy not in (None, 0):
        pct = (spread / buy) * 100

    return buy, sell, spread, pct


def read_history(path):

    path = Path(path)

    if not path.exists():
        return []

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _style_header(ws, row, start_col, end_col):

    for col in range(start_col, end_col + 1):

        cell = ws.cell(row, col)

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def _set_widths(ws, widths):

    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _apply_table(ws, start_row, end_row, end_col, name):

    if end_row <= start_row:
        return

    ref = (
        f"A{start_row}:"
        f"{get_column_letter(end_col)}{end_row}"
    )

    table = Table(
        displayName=name,
        ref=ref,
    )

    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False,
    )

    ws.add_table(table)


# ============================================================
# TARİH ETİKETİ
# ============================================================

def _date_label(dt, include_time=False):

    if dt.weekday() == 5:
        suffix = " Cmt"
    elif dt.weekday() == 6:
        suffix = " Paz"
    else:
        suffix = ""

    result = dt.strftime("%d.%m") + suffix

    if include_time:
        result += " " + dt.strftime("%H:%M")

    return result


def _long_date_label(dt):

    if dt.weekday() == 5:
        suffix = " - Cumartesi"
    elif dt.weekday() == 6:
        suffix = " - Pazar"
    else:
        suffix = ""

    return dt.strftime("%d.%m.%Y") + suffix


# ============================================================
# GÜNLÜK İLK RUN
# ============================================================

def _daily_history(history):

    first_runs = {}

    for row in history:

        raw = row.get("run_at")
        dt = _parse_dt(raw)

        if not raw or not dt:
            continue

        day = dt.date()

        old_raw = first_runs.get(day)

        if old_raw is None:
            first_runs[day] = raw
            continue

        old_dt = _parse_dt(old_raw)

        if old_dt is None or dt < old_dt:
            first_runs[day] = raw

    selected = set(first_runs.values())

    return [
        row
        for row in history
        if row.get("run_at") in selected
    ]


# ============================================================
# SON RUN
# ============================================================

def _latest_run_rows(history):

    valid = []

    for row in history:

        dt = _parse_dt(row.get("run_at"))

        if dt:
            valid.append((dt, row))

    if not valid:
        return None, []

    latest_dt = max(x[0] for x in valid)

    rows = [
        row
        for dt, row in valid
        if dt == latest_dt
    ]

    return rows[0].get("run_at"), rows


# ============================================================
# PROVIDER MAP
# ============================================================

def _provider_map(rows):

    result = {}

    for row in rows:

        provider = (row.get("provider") or "").strip()
        code = (row.get("code") or "").strip().upper()

        if not provider or not code:
            continue

        result.setdefault(provider, {})[code] = row

    return result


# ============================================================
# GUNCEL KURLAR
# ============================================================

def _build_current_sheet(wb, latest_run_at, latest_rows):

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

    _style_header(ws, 1, 1, 15)

    for col in range(4, 8):
        ws.cell(1, col).fill = USD_HEADER_FILL

    for col in range(8, 12):
        ws.cell(1, col).fill = EUR_HEADER_FILL

    for col in range(12, 16):
        ws.cell(1, col).fill = XAU_HEADER_FILL

    run_dt = _parse_dt(latest_run_at)

    providers = _provider_map(latest_rows)

    for row_no, provider in enumerate(
        sorted(providers, key=str.casefold),
        start=2,
    ):

        data = providers[provider]

        if run_dt:
            ws.cell(row_no, 1, run_dt.date())
            ws.cell(row_no, 2, run_dt.time())

        provider_cell = ws.cell(row_no, 3, provider)

        provider_cell.fill = _provider_fill(provider)
        provider_cell.font = Font(bold=True)

        layout = {
            "USD": (4, 5, 6, 7),
            "EUR": (8, 9, 10, 11),
            "XAU": (12, 13, 14, 15),
        }

        for code, columns in layout.items():

            for col in columns:
                ws.cell(row_no, col).fill = _product_fill(code)

            item = data.get(code)

            if not item:
                continue

            buy, sell, spread, pct = _get_spread_values(item)

            if buy is not None:
                ws.cell(row_no, columns[0], buy)

            if sell is not None:
                ws.cell(row_no, columns[1], sell)

            if spread is not None:
                ws.cell(row_no, columns[2], spread)

            if pct is not None:
                ws.cell(row_no, columns[3], pct / 100)

        ws.cell(row_no, 1).number_format = "dd.mm.yyyy"
        ws.cell(row_no, 2).number_format = "hh:mm:ss"

        for col in (4, 5, 6, 8, 9, 10, 12, 13, 14):
            ws.cell(row_no, col).number_format = "#,##0.0000"

        for col in (7, 11, 15):
            ws.cell(row_no, col).number_format = "0.00%"

    if ws.max_row >= 2:

        _apply_table(
            ws,
            1,
            ws.max_row,
            15,
            "GuncelKurlarTable",
        )

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


# ============================================================
# GECMIS
# ============================================================

def _history_sort_key(row):

    dt = _parse_dt(row.get("run_at")) or datetime.min
    provider = (row.get("provider") or "").casefold()
    code = (row.get("code") or "").upper()

    return (
        dt,
        provider,
        PRODUCT_ORDER.get(code, 99),
    )


def _build_history_sheet(wb, history):

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

    for row_no, item in enumerate(
        sorted(history, key=_history_sort_key),
        start=2,
    ):

        dt = _parse_dt(item.get("run_at"))
        scraped = _parse_dt(item.get("scraped_at"))

        code = (item.get("code") or "").upper()

        buy, sell, spread, pct = _get_spread_values(item)

        site_spread = _to_float(item.get("site_spread"))
        site_pct = _to_float(item.get("site_spread_pct"))

        provider = item.get("provider", "")

        if dt:
            ws.cell(row_no, 1, dt.date())
            ws.cell(row_no, 2, dt.time())

        ws.cell(row_no, 3, provider)

        ws.cell(row_no, 3).fill = _provider_fill(provider)
        ws.cell(row_no, 3).font = Font(bold=True)

        ws.cell(
            row_no,
            4,
            item.get("product") or PRODUCT_NAMES.get(code, code),
        )

        for col in range(4, 9):
            ws.cell(row_no, col).fill = _product_fill(code)

        if buy is not None:
            ws.cell(row_no, 5, buy)

        if sell is not None:
            ws.cell(row_no, 6, sell)

        if spread is not None:
            ws.cell(row_no, 7, spread)

        if pct is not None:
            ws.cell(row_no, 8, pct / 100)

        status = item.get("status", "")

        ws.cell(
            row_no,
            9,
            STATUS_LABELS.get(status, status),
        )

        ws.cell(
            row_no,
            10,
            item.get("source_url", ""),
        )

        if site_spread is not None:
            ws.cell(row_no, 11, site_spread)

        if site_pct is not None:
            ws.cell(row_no, 12, site_pct / 100)

        ws.cell(
            row_no,
            13,
            item.get("note", ""),
        )

        if scraped:
            ws.cell(row_no, 14, scraped.time())

        if dt and dt.weekday() in (5, 6):

            ws.cell(row_no, 1).fill = WEEKEND_FILL

            ws.cell(row_no, 1).font = Font(
                bold=True,
                color="9C6500",
            )

        ws.cell(row_no, 1).number_format = "dd.mm.yyyy"
        ws.cell(row_no, 2).number_format = "hh:mm:ss"
        ws.cell(row_no, 14).number_format = "hh:mm:ss"

        for col in (5, 6, 7, 11):
            ws.cell(row_no, col).number_format = "#,##0.0000"

        for col in (8, 12):
            ws.cell(row_no, col).number_format = "0.00%"

        if ws.cell(row_no, 10).value:
            ws.cell(row_no, 10).hyperlink = ws.cell(row_no, 10).value
            ws.cell(row_no, 10).style = "Hyperlink"

        if status == "ERROR":
            ws.cell(row_no, 9).fill = ERROR_FILL
        elif status == "KONTROL":
            ws.cell(row_no, 9).fill = CONTROL_FILL
        else:
            ws.cell(row_no, 9).fill = OK_FILL

        ws.cell(row_no, 9).font = Font(bold=True)

    if ws.max_row >= 2:
        _apply_table(
            ws,
            1,
            ws.max_row,
            len(headers),
            "GecmisTable",
        )

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
            "I": 20,
            "J": 45,
            "K": 17,
            "L": 18,
            "M": 60,
            "N": 22,
        },
    )


# ============================================================
# TREND VERİLERİ
# ============================================================

def _build_bank_trends(history):

    data = {}

    for row in history:

        bank = (row.get("provider") or "").strip()
        code = (row.get("code") or "").upper()

        if bank not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if row.get("status") == "ERROR":
            continue

        dt = _parse_dt(row.get("run_at"))

        if not dt:
            continue

        _, _, _, pct = _get_spread_values(row)

        if pct is None:
            continue

        key = row.get("run_at")

        bucket = data.setdefault(
            key,
            {
                "dt": dt,
                "banks": {},
            },
        )

        bucket["banks"].setdefault(
            bank,
            {},
        )[code] = pct / 100

    return sorted(
        data.values(),
        key=lambda x: x["dt"],
    )


# ============================================================
# AYLIK ORTALAMA
# ============================================================

def _monthly_averages(history):

    buckets = {}

    for row in history:

        bank = (row.get("provider") or "").strip()
        code = (row.get("code") or "").upper()

        if bank not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if row.get("status") == "ERROR":
            continue

        dt = _parse_dt(row.get("run_at"))

        if not dt:
            continue

        _, _, _, pct = _get_spread_values(row)

        if pct is None:
            continue

        key = (
            dt.year,
            dt.month,
            bank,
            code,
        )

        buckets.setdefault(key, []).append(
            pct / 100
        )

    result = {}

    for (
        year,
        month,
        bank,
        code,
    ), values in buckets.items():

        result.setdefault(
            (year, month),
            {},
        ).setdefault(
            bank,
            {},
        )[code] = sum(values) / len(values)

    return result


# ============================================================
# GRAFİK STİLİ
# ============================================================

def _set_series_color(
    series,
    color,
    marker="circle",
    size=6,
):
    try:
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 22000
        series.marker.symbol = marker
        series.marker.size = size
        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color
        series.smooth = False
    except Exception:
        pass


def _style_chart(chart, title, legend=True):
    chart.title = title
    chart.width = 16.0
    chart.height = 8.0

    # Y ekseni: sayısal yüzde ölçeği mutlaka görünür.
    chart.y_axis.title = "Makas %"
    chart.y_axis.numFmt = "0.00%"
    chart.y_axis.delete = False
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.majorTickMark = "out"
    chart.y_axis.minorTickMark = "none"
    chart.y_axis.majorGridlines = None

    # X ekseni: helper sayfasındaki metin tarihleri kategori etiketi olarak göster.
    chart.x_axis.delete = False
    chart.x_axis.tickLblPos = "low"
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.minorTickMark = "none"
    chart.x_axis.tickLblSkip = 1
    chart.x_axis.tickMarkSkip = 1
    chart.x_axis.majorGridlines = None

    # İki ekseni açıkça birbirine bağla. Excel'in eksen etiketlerini
    # gizlemesine yol açan varsayılan/bozuk cross ayarlarını engeller.
    try:
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
    except Exception:
        pass

    try:
        chart.visible_cells_only = False
    except Exception:
        pass

    try:
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot=-2700000),
            p=[
                Paragraph(
                    pPr=ParagraphProperties(
                        defRPr=CharacterProperties(sz=700)
                    ),
                    endParaRPr=CharacterProperties(sz=700),
                )
            ],
        )
        chart.y_axis.txPr = RichText(
            bodyPr=RichTextProperties(),
            p=[
                Paragraph(
                    pPr=ParagraphProperties(
                        defRPr=CharacterProperties(sz=700)
                    ),
                    endParaRPr=CharacterProperties(sz=700),
                )
            ],
        )
    except Exception:
        pass

    if legend:
        chart.legend.position = "b"
        chart.legend.overlay = False
    else:
        chart.legend = None


# ============================================================
# CACHE
# ============================================================

def _cache_line_chart(chart, categories, series_data):

    try:

        for series, (title, values) in zip(
            chart.series,
            series_data,
        ):

            points = []

            for i, value in enumerate(values):

                if value is None:
                    continue

                points.append(
                    NumVal(
                        idx=i,
                        v=float(value),
                    )
                )

            if series.val and series.val.numRef:

                series.val.numRef.numCache = NumData(
                    formatCode="0.00%",
                    ptCount=len(values),
                    pt=points,
                )

            if series.cat:

                formula = None

                if series.cat.numRef:
                    formula = series.cat.numRef.f

                elif series.cat.strRef:
                    formula = series.cat.strRef.f

                if formula:

                    series.cat.numRef = None

                    series.cat.strRef = StrRef(
                        f=formula,
                        strCache=StrData(
                            ptCount=len(categories),
                            pt=[
                                StrVal(
                                    idx=i,
                                    v=str(v),
                                )
                                for i, v
                                in enumerate(categories)
                            ],
                        ),
                    )

            if series.tx and series.tx.strRef:

                series.tx.strRef.strCache = StrData(
                    ptCount=1,
                    pt=[
                        StrVal(
                            idx=0,
                            v=title,
                        )
                    ],
                )

    except Exception:
        pass


# ============================================================
# GRAFİK VERİLERİ SAYFASI
#
# OZET'TE ARTIK AA-AU TEKNİK VERİLER YOK.
#
# Helper verileri ayrı sayfada.
#
# Sayfayı "veryHidden" yapmıyoruz.
# Normal hidden yapıyoruz; Excel grafik referansları açısından
# daha güvenli.
# ============================================================

def _build_chart_data_sheet(wb, trends, monthly):

    ws = wb.create_sheet("GRAFIK_VERILERI")

    # --------------------------------------------------------
    # GÜNLÜK
    # --------------------------------------------------------

    ws["A1"] = "Tarih"

    col = 2

    daily_columns = {}

    for code in ("USD", "EUR", "XAU"):

        daily_columns[code] = {}

        for bank in TARGET_BANKS:

            daily_columns[code][bank] = col

            ws.cell(
                1,
                col,
                f"{code} - {bank}",
            )

            col += 1

    for row_no, run in enumerate(
        trends,
        start=2,
    ):

        ws.cell(
            row_no,
            1,
            _date_label(run["dt"]),
        )

        for code in ("USD", "EUR", "XAU"):

            for bank in TARGET_BANKS:

                value = (
                    run["banks"]
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    ws.cell(
                        row_no,
                        daily_columns[code][bank],
                        value,
                    ).number_format = "0.00%"

    # --------------------------------------------------------
    # AYLIK
    # --------------------------------------------------------

    months = sorted(monthly.keys())

    monthly_start_col = 20

    ws.cell(
        1,
        monthly_start_col,
        "Ay",
    )

    monthly_columns = {}

    col = monthly_start_col + 1

    for code in ("USD", "EUR", "XAU"):

        monthly_columns[code] = {}

        for bank in TARGET_BANKS:

            monthly_columns[code][bank] = col

            ws.cell(
                1,
                col,
                f"{code} - {bank}",
            )

            col += 1

    for row_no, month_key in enumerate(
        months,
        start=2,
    ):

        year, month = month_key

        ws.cell(
            row_no,
            monthly_start_col,
            f"{MONTH_NAMES[month]} {year}",
        )

        for code in ("USD", "EUR", "XAU"):

            for bank in TARGET_BANKS:

                value = (
                    monthly
                    .get(month_key, {})
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    ws.cell(
                        row_no,
                        monthly_columns[code][bank],
                        value,
                    ).number_format = "0.00%"

    return (
        ws,
        daily_columns,
        monthly_columns,
        monthly_start_col,
        months,
    )


# ============================================================
# OZET
# ============================================================

def _build_summary_sheet(wb, latest_run_at, latest_rows, history):
    ws = wb.create_sheet("OZET")
    ws.sheet_view.showGridLines = False

    trends = _build_bank_trends(history)
    monthly = _monthly_averages(history)
    data_ws, daily_columns, monthly_columns, monthly_date_col, months = _build_chart_data_sheet(wb, trends, monthly)

    ws.merge_cells("A1:Q1")
    ws["A1"] = "Döviz ve Altın Kur Takip Özeti"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # SON VERİ ÇEKİMİ - SOL
    ws.merge_cells("A3:D3")
    ws["A3"] = "SON VERİ ÇEKİMİ"
    ws["A3"].fill = TITLE_FILL
    ws["A3"].font = Font(bold=True, color="1F4E78")
    ws["A3"].alignment = Alignment(horizontal="center")
    latest_dt = _parse_dt(latest_run_at)
    providers = {(r.get("provider") or "").strip() for r in latest_rows if (r.get("provider") or "").strip()}
    info = [
        ("Veri tarihi", latest_dt.date() if latest_dt else ""),
        ("Çekim saati", latest_dt.time() if latest_dt else ""),
        ("Sağlayıcı sayısı", len(providers)),
        ("Ürün kaydı", len(latest_rows)),
        ("Hatalı kayıt", sum(1 for r in latest_rows if r.get("status") == "ERROR")),
        ("Kontrol gereken", sum(1 for r in latest_rows if r.get("status") == "KONTROL")),
    ]
    for row_no, (label, value) in enumerate(info, start=4):
        ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=2)
        ws.merge_cells(start_row=row_no, start_column=3, end_row=row_no, end_column=4)
        ws.cell(row_no, 1, label)
        ws.cell(row_no, 3, value)
        ws.cell(row_no, 1).font = Font(bold=True)
        ws.cell(row_no, 1).fill = PatternFill("solid", fgColor="EAF2F8")
        ws.cell(row_no, 1).alignment = Alignment(vertical="center")
        ws.cell(row_no, 3).alignment = Alignment(horizontal="center")
        for c in range(1, 5):
            ws.cell(row_no, c).border = THIN_BORDER
    ws["C4"].number_format = "dd.mm.yyyy"
    ws["C5"].number_format = "hh:mm:ss"

    # AYLIK ORTALAMA - ORTA
    ws.merge_cells("F3:J3")
    ws["F3"] = "AYLIK ORTALAMA MAKAS %"
    ws["F3"].fill = TITLE_FILL
    ws["F3"].font = Font(bold=True, color="1F4E78")
    ws["F3"].alignment = Alignment(horizontal="center")
    for col, value in enumerate(["Ay", "Banka", "Dolar", "Euro", "Gram Altın"], start=6):
        ws.cell(4, col, value)
    _style_header(ws, 4, 6, 10)
    ws["H4"].fill = USD_HEADER_FILL
    ws["I4"].fill = EUR_HEADER_FILL
    ws["J4"].fill = XAU_HEADER_FILL
    monthly_row = 5
    for year, month in months:
        label = f"{MONTH_NAMES[month]} {year}"
        for bank in TARGET_BANKS:
            ws.cell(monthly_row, 6, label)
            ws.cell(monthly_row, 7, bank)
            ws.cell(monthly_row, 7).fill = _provider_fill(bank)
            ws.cell(monthly_row, 7).font = Font(bold=True)
            for col, code in ((8, "USD"), (9, "EUR"), (10, "XAU")):
                value = monthly.get((year, month), {}).get(bank, {}).get(code)
                ws.cell(monthly_row, col).fill = _product_fill(code)
                if value is not None:
                    ws.cell(monthly_row, col, value)
                    ws.cell(monthly_row, col).number_format = "0.00%"
            for col in range(6, 11):
                ws.cell(monthly_row, col).border = THIN_BORDER
            monthly_row += 1
    monthly_last_row = max(monthly_row - 1, 10)

    # EN DÜŞÜK MAKAS - SAĞ
    ws.merge_cells("L3:Q3")
    ws["L3"] = "5 BANKADA EN DÜŞÜK GÜNCEL MAKAS"
    ws["L3"].fill = TITLE_FILL
    ws["L3"].font = Font(bold=True, color="1F4E78")
    ws["L3"].alignment = Alignment(horizontal="center")
    best_headers = ["Ürün", "Banka Sayısı", "En Düşük %", "En Avantajlı", "Alış", "Satış"]
    for col, value in enumerate(best_headers, start=12):
        ws.cell(4, col, value)
    _style_header(ws, 4, 12, 17)
    best_row = 5
    for code in ("USD", "EUR", "XAU"):
        candidates = []
        banks_seen = set()
        for item in latest_rows:
            bank = (item.get("provider") or "").strip()
            item_code = (item.get("code") or "").upper()
            if bank not in TARGET_BANKS or item_code != code or item.get("status") == "ERROR":
                continue
            buy, sell, _, pct = _get_spread_values(item)
            if buy is not None and sell is not None and pct is not None:
                banks_seen.add(bank)
                candidates.append((pct, bank, buy, sell))
        best = min(candidates, key=lambda x: x[0]) if candidates else None
        ws.cell(best_row, 12, PRODUCT_NAMES[code])
        ws.cell(best_row, 12).fill = _product_fill(code)
        ws.cell(best_row, 12).font = Font(bold=True)
        ws.cell(best_row, 13, len(banks_seen))
        if best:
            pct, bank, buy, sell = best
            ws.cell(best_row, 14, pct / 100)
            ws.cell(best_row, 14).number_format = "0.00%"
            ws.cell(best_row, 15, bank)
            ws.cell(best_row, 15).fill = _provider_fill(bank)
            ws.cell(best_row, 15).font = Font(bold=True)
            ws.cell(best_row, 16, buy)
            ws.cell(best_row, 17, sell)
            ws.cell(best_row, 16).number_format = "#,##0.0000"
            ws.cell(best_row, 17).number_format = "#,##0.0000"
        for col in range(12, 18):
            ws.cell(best_row, col).border = THIN_BORDER
        best_row += 1

    # Grafikler tablolar bittikten sonra başlar. Günlük detay tablosu YOK.
    graph_start = monthly_last_row + 4
    GRAPH_COLS = ["A", "J", "S"]
    GRAPH_ROW_GAP = 24
    daily_end_row = len(trends) + 1
    category_labels = [_date_label(run["dt"]) for run in trends]

    def set_bounds(chart, values, ratio=0.15):
        valid = [v for v in values if v is not None]
        if not valid:
            return
        minimum, maximum = min(valid), max(valid)
        padding = max((maximum - minimum) * ratio, maximum * 0.02, 0.00005)
        chart.y_axis.scaling.min = max(0, minimum - padding)
        chart.y_axis.scaling.max = maximum + padding

    if trends:
        daily_categories = Reference(data_ws, min_col=1, min_row=2, max_row=daily_end_row)
        for product_index, code in enumerate(("USD", "EUR", "XAU")):
            chart = LineChart()
            _style_chart(chart, f"{PRODUCT_NAMES[code]} - 5 Banka Makas %", legend=True)
            cached, all_values = [], []
            for bank in TARGET_BANKS:
                col = daily_columns[code][bank]
                data = Reference(data_ws, min_col=col, max_col=col, min_row=1, max_row=daily_end_row)
                chart.add_data(data, titles_from_data=True)
                values = [run["banks"].get(bank, {}).get(code) for run in trends]
                cached.append((bank, values))
                all_values.extend(v for v in values if v is not None)
            chart.set_categories(daily_categories)
            for index, series in enumerate(chart.series):
                bank = TARGET_BANKS[index]
                _set_series_color(series, BANK_LINE_COLORS[bank], marker="circle", size=6)
            _cache_line_chart(chart, category_labels, cached)
            set_bounds(chart, all_values)
            chart.dLbls = None
            ws.add_chart(chart, f"{GRAPH_COLS[product_index]}{graph_start}")

    monthly_graph_row = graph_start + GRAPH_ROW_GAP
    monthly_end_row = len(months) + 1
    if months:
        monthly_categories = Reference(data_ws, min_col=monthly_date_col, min_row=2, max_row=monthly_end_row)
        monthly_labels = [f"{MONTH_NAMES[m]} {y}" for y, m in months]
        for product_index, code in enumerate(("USD", "EUR", "XAU")):
            chart = LineChart()
            _style_chart(chart, f"{PRODUCT_NAMES[code]} - Aylık Ortalama Makas %", legend=True)
            cached, all_values = [], []
            for bank in TARGET_BANKS:
                col = monthly_columns[code][bank]
                data = Reference(data_ws, min_col=col, max_col=col, min_row=1, max_row=monthly_end_row)
                chart.add_data(data, titles_from_data=True)
                values = [monthly.get(month_key, {}).get(bank, {}).get(code) for month_key in months]
                cached.append((bank, values))
                all_values.extend(v for v in values if v is not None)
            chart.set_categories(monthly_categories)
            for index, series in enumerate(chart.series):
                bank = TARGET_BANKS[index]
                _set_series_color(series, BANK_LINE_COLORS[bank], marker="circle", size=7)
            _cache_line_chart(chart, monthly_labels, cached)
            set_bounds(chart, all_values)
            chart.dLbls = None
            ws.add_chart(chart, f"{GRAPH_COLS[product_index]}{monthly_graph_row}")

    individual_start = monthly_graph_row + GRAPH_ROW_GAP
    if trends:
        daily_categories = Reference(data_ws, min_col=1, min_row=2, max_row=daily_end_row)
        for bank_index, bank in enumerate(TARGET_BANKS):
            bank_row = individual_start + bank_index * GRAPH_ROW_GAP
            for product_index, code in enumerate(("USD", "EUR", "XAU")):
                chart = LineChart()
                _style_chart(chart, f"{bank} - {PRODUCT_NAMES[code]} Makas %", legend=False)
                try:
                    bg = BANK_CHART_COLORS.get(bank, DEFAULT_BANK_CHART_COLOR)
                    chart.graphical_properties = GraphicalProperties(solidFill=bg)
                    chart.plot_area.graphicalProperties = GraphicalProperties(solidFill=bg)
                except Exception:
                    pass
                col = daily_columns[code][bank]
                data = Reference(data_ws, min_col=col, max_col=col, min_row=1, max_row=daily_end_row)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(daily_categories)
                if chart.series:
                    marker = {"USD": "circle", "EUR": "square", "XAU": "triangle"}[code]
                    _set_series_color(chart.series[0], PRODUCT_LINE_COLORS[code], marker=marker, size=6)
                values = [run["banks"].get(bank, {}).get(code) for run in trends]
                _cache_line_chart(chart, category_labels, [(PRODUCT_NAMES[code], values)])
                set_bounds(chart, values, ratio=0.20)

                # Tekil banka grafiklerinde noktanın sayısal Makas % değeri.
                chart.dLbls = DataLabelList()
                chart.dLbls.showVal = True
                chart.dLbls.showSerName = False
                chart.dLbls.showCatName = False
                chart.dLbls.showLegendKey = False
                chart.dLbls.numFmt = "0.00%"
                chart.dLbls.dLblPos = "t"

                # Tarih ve Y ekseni etiketlerini ayrıca zorla görünür tut.
                chart.x_axis.delete = False
                chart.x_axis.tickLblPos = "low"
                chart.x_axis.tickLblSkip = 1
                chart.y_axis.delete = False
                chart.y_axis.tickLblPos = "nextTo"

                ws.add_chart(chart, f"{GRAPH_COLS[product_index]}{bank_row}")

    _set_widths(ws, {
        "A": 15, "B": 15, "C": 15, "D": 15, "E": 3,
        "F": 17, "G": 20, "H": 14, "I": 14, "J": 17, "K": 3,
        "L": 17, "M": 14, "N": 15, "O": 19, "P": 16, "Q": 16, "R": 3,
        "S": 12, "T": 12, "U": 12, "V": 12, "W": 12, "X": 12, "Y": 12, "Z": 12,
    })

    final_row = individual_start + len(TARGET_BANKS) * GRAPH_ROW_GAP + 5
    for row_no in range(graph_start, final_row):
        ws.row_dimensions[row_no].height = 20

    data_ws.sheet_state = "hidden"


# ============================================================
# ANA FONKSİYON
# ============================================================

def build_excel(
    history_path: str | Path,
    output_path: str | Path,
) -> None:

    raw_history = read_history(
        history_path
    )

    if not raw_history:
        raise RuntimeError(
            "Excel oluşturmak için geçmiş veri bulunamadı."
        )

    latest_run_at, latest_rows = _latest_run_rows(
        raw_history
    )

    if not latest_rows:
        raise RuntimeError(
            "Son çekime ait veri bulunamadı."
        )

    # Grafikler / geçmiş / aylık ortalama:
    # aynı günün yalnızca ilk run'ı.
    daily_history = _daily_history(
        raw_history
    )

    if not daily_history:
        raise RuntimeError(
            "Günlük geçmiş veri bulunamadı."
        )

    wb = Workbook()

    wb.remove(
        wb.active
    )

    # --------------------------------------------------------
    # 1
    # --------------------------------------------------------

    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )

    # --------------------------------------------------------
    # 2
    # --------------------------------------------------------

    _build_history_sheet(
        wb,
        daily_history,
    )

    # --------------------------------------------------------
    # 3
    # --------------------------------------------------------

    _build_summary_sheet(
        wb,
        latest_run_at,
        latest_rows,
        daily_history,
    )

    wb.active = 0

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    wb.save(
        output_path
    )


# ============================================================
# ESKİ main.py İSİMLERİYLE UYUMLULUK
# ============================================================

def write_excel(
    history_path,
    output_path,
):
    return build_excel(
        history_path,
        output_path,
    )


def create_excel(
    history_path,
    output_path,
):
    return build_excel(
        history_path,
        output_path,
    )


def create_excel_report(
    history_path,
    output_path,
):
    return build_excel(
        history_path,
        output_path,
    )
