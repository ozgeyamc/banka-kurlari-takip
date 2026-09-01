from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.data_source import (
    NumData,
    NumVal,
    StrData,
    StrRef,
    StrVal,
)
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


# ============================================================
# GENEL STİLLER
# ============================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FILL = PatternFill("solid", fgColor="D9EAF7")

THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)

OK_FILL = PatternFill("solid", fgColor="E2F0D9")
CONTROL_FILL = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL = PatternFill("solid", fgColor="FCE4D6")
BEST_FILL = PatternFill("solid", fgColor="E2F0D9")
BEST_FONT = Font(color="006100", bold=True)

WEEKEND_FILL = PatternFill("solid", fgColor="FFE699")

STATUS_LABELS = {
    "OK": "DOĞRU",
    "KONTROL": "KONTROL GEREKLİ",
    "ERROR": "HATA",
}


# ============================================================
# ÜRÜN RENKLERİ
# ============================================================

USD_FILL = PatternFill("solid", fgColor="DDEBF7")
EUR_FILL = PatternFill("solid", fgColor="FCE4D6")
XAU_FILL = PatternFill("solid", fgColor="FFF2CC")

USD_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
EUR_HEADER_FILL = PatternFill("solid", fgColor="ED7D31")
XAU_HEADER_FILL = PatternFill("solid", fgColor="BF9000")

PRODUCT_LINE_COLORS = {
    "USD": "4472C4",
    "EUR": "ED7D31",
    "XAU": "FFC000",
}

WEEKEND_MARKER_COLOR = "FF0000"


# ============================================================
# SAĞLAYICI RENKLERİ
# ============================================================

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
    "Yapıkredi": "E4D5F7",
    "Ziraat Bankası": "DFF1E3",
    "Ziraat Dinamik": "DBCAEC",
    "Ziraat Katılım": "F1EDDF",
}

DEFAULT_PROVIDER_COLOR = "E8EDF3"


# ============================================================
# BANKA GRAFİK ARKA PLANLARI
# ============================================================

BANK_CHART_COLORS = {
    "Akbank": "FDE2E2",
    "Garanti BBVA": "E2F0D9",

    # Yapıkredi artık belirgin mor/lila
    "Yapıkredi": "E4D5F7",

    "Ziraat Bankası": "FFF2CC",
    "İş Bankası": "DDEBF7",
}

DEFAULT_BANK_CHART_COLOR = "E8EDF3"


# ============================================================
# BANKA ÇİZGİ RENKLERİ
# ============================================================

BANK_LINE_COLORS = {
    "Garanti BBVA": "70AD47",   # yeşil
    "Akbank": "C00000",         # kırmızı
    "Yapıkredi": "7030A0",      # mor
    "Ziraat Bankası": "FFC000", # sarı
    "İş Bankası": "4472C4",     # mavi
}


# ============================================================
# ÜRÜNLER
# ============================================================

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

TARGET_BANKS = [
    "Garanti BBVA",
    "Akbank",
    "Yapıkredi",
    "Ziraat Bankası",
    "İş Bankası",
]

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
# YARDIMCI FONKSİYONLAR
# ============================================================

def _provider_fill(provider: str | None):
    color = PROVIDER_COLORS.get(
        (provider or "").strip(),
        DEFAULT_PROVIDER_COLOR,
    )
    return PatternFill("solid", fgColor=color)


def _product_fill(code: str | None):
    code = (code or "").strip().upper()

    if code == "USD":
        return USD_FILL

    if code == "EUR":
        return EUR_FILL

    if code == "XAU":
        return XAU_FILL

    return PatternFill(fill_type=None)


def _parse_dt(value: str | None) -> datetime | None:
    if not value:
        return None

    try:
        return datetime.fromisoformat(value)
    except (ValueError, TypeError):
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

    with path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        return list(csv.DictReader(handle))


def _style_header(ws, row: int, start_col: int, end_col: int):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def _apply_table(
    ws,
    start_row: int,
    end_row: int,
    end_col: int,
    name: str,
):
    if end_row <= start_row:
        return

    from openpyxl.utils import get_column_letter

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


def _set_widths(ws, widths):
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def _status_fill(status):
    if status == "ERROR":
        return ERROR_FILL

    if status == "KONTROL":
        return CONTROL_FILL

    return OK_FILL


def _display_status(status):
    raw = (status or "").strip()
    return STATUS_LABELS.get(raw, raw)


# ============================================================
# HER GÜNÜN İLK RUN'I
# ============================================================

def _daily_history(history: list[dict]) -> list[dict]:

    first_run_string_by_day = {}

    for row in history:
        raw = row.get("run_at")
        dt = _parse_dt(raw)

        if not dt or not raw:
            continue

        day = dt.date()

        current = first_run_string_by_day.get(day)

        if current is None:
            first_run_string_by_day[day] = raw
            continue

        current_dt = _parse_dt(current)

        if current_dt is None or dt < current_dt:
            first_run_string_by_day[day] = raw

    selected = set(
        first_run_string_by_day.values()
    )

    return [
        row
        for row in history
        if row.get("run_at") in selected
    ]


# ============================================================
# EN SON RUN
# ============================================================

def _latest_run_rows(history):
    valid = [
        row
        for row in history
        if _parse_dt(row.get("run_at"))
    ]

    if not valid:
        return None, []

    latest = max(
        valid,
        key=lambda row: _parse_dt(
            row.get("run_at")
        ),
    )

    latest_run_at = latest["run_at"]

    return (
        latest_run_at,
        [
            row
            for row in history
            if row.get("run_at") == latest_run_at
        ],
    )


def _provider_map(rows: Iterable[dict]):
    result = {}

    for row in rows:
        provider = (
            row.get("provider", "")
            .strip()
        )

        code = (
            row.get("code", "")
            .strip()
        )

        if not provider or not code:
            continue

        result.setdefault(
            provider,
            {},
        )[code] = row

    return result


def _history_sort_key(row):
    provider = (
        row.get("provider")
        or ""
    ).strip().casefold()

    code = (
        row.get("code")
        or ""
    ).strip()

    run_dt = (
        _parse_dt(row.get("run_at"))
        or datetime.min
    )

    return (
        provider,
        PRODUCT_ORDER.get(code, 99),
        run_dt,
    )


# ============================================================
# GRAFİK CACHE
# ============================================================

def _cache_line_chart(
    chart,
    categories,
    series_values,
):

    for series, (title, values) in zip(
        chart.series,
        series_values,
    ):

        numeric_points = [
            NumVal(
                idx=index,
                v=float(value),
            )
            for index, value in enumerate(values)
            if value is not None
        ]

        if (
            series.val is not None
            and series.val.numRef is not None
        ):
            series.val.numRef.numCache = NumData(
                formatCode="0.000%",
                ptCount=len(values),
                pt=numeric_points,
            )

        if series.cat is not None:
            category_formula = None

            if series.cat.numRef is not None:
                category_formula = series.cat.numRef.f

            elif series.cat.strRef is not None:
                category_formula = series.cat.strRef.f

            series.cat.numRef = None

            series.cat.strRef = StrRef(
                f=category_formula,
                strCache=StrData(
                    ptCount=len(categories),
                    pt=[
                        StrVal(
                            idx=index,
                            v=str(value),
                        )
                        for index, value
                        in enumerate(categories)
                    ],
                ),
            )

        if (
            series.tx is not None
            and series.tx.strRef is not None
        ):
            series.tx.strRef.strCache = StrData(
                ptCount=1,
                pt=[
                    StrVal(
                        idx=0,
                        v=title,
                    )
                ],
            )


# ============================================================
# HAFTA SONU MARKER
# ============================================================

def _apply_weekend_markers(series, dates):
    """
    Cumartesi ve Pazar noktalarını kırmızı ve daha büyük gösterir.
    """

    points = []

    for index, dt in enumerate(dates):

        if dt.weekday() not in (5, 6):
            continue

        point = DataPoint(idx=index)

        try:
            point.marker.symbol = "diamond"
            point.marker.size = 9
            point.marker.graphicalProperties.solidFill = (
                WEEKEND_MARKER_COLOR
            )
            point.marker.graphicalProperties.line.solidFill = (
                WEEKEND_MARKER_COLOR
            )
        except Exception:
            pass

        points.append(point)

    if points:
        try:
            series.dPt = points
        except Exception:
            pass


# ============================================================
# GÜNCEL KURLAR
# ============================================================

def _build_current_sheet(
    wb,
    latest_run_at,
    latest_rows,
):

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

    _style_header(
        ws,
        1,
        1,
        len(headers),
    )

    for col in range(4, 8):
        ws.cell(1, col).fill = USD_HEADER_FILL

    for col in range(8, 12):
        ws.cell(1, col).fill = EUR_HEADER_FILL

    for col in range(12, 16):
        ws.cell(1, col).fill = XAU_HEADER_FILL

    provider_map = _provider_map(latest_rows)

    run_dt = _parse_dt(latest_run_at)

    for excel_row, provider in enumerate(
        sorted(provider_map, key=str.casefold),
        start=2,
    ):

        row_map = provider_map[provider]

        ws.cell(
            excel_row,
            1,
            run_dt.date() if run_dt else "",
        )

        ws.cell(
            excel_row,
            2,
            (
                run_dt.time().replace(tzinfo=None)
                if run_dt
                else ""
            ),
        )

        provider_cell = ws.cell(
            excel_row,
            3,
            provider,
        )

        provider_cell.fill = _provider_fill(provider)
        provider_cell.font = Font(bold=True)

        layout = {
            "USD": (4, 5, 6, 7),
            "EUR": (8, 9, 10, 11),
            "XAU": (12, 13, 14, 15),
        }

        for code, cols in layout.items():

            buy_col, sell_col, spread_col, pct_col = cols

            fill = _product_fill(code)

            for col in cols:
                ws.cell(
                    excel_row,
                    col,
                ).fill = fill

            item = row_map.get(code)

            if not item:
                continue

            buy = _to_float(item.get("buy"))
            sell = _to_float(item.get("sell"))
            spread = _to_float(item.get("spread"))
            spread_pct = _to_float(
                item.get("spread_pct")
            )

            if buy is not None:
                ws.cell(
                    excel_row,
                    buy_col,
                    buy,
                )

            if sell is not None:
                ws.cell(
                    excel_row,
                    sell_col,
                    sell,
                )

            if (
                spread is None
                and buy is not None
                and sell is not None
            ):
                spread = sell - buy

            if (
                spread_pct is None
                and spread is not None
                and buy not in (None, 0)
            ):
                spread_pct = (
                    spread / buy
                ) * 100

            if spread is not None:
                ws.cell(
                    excel_row,
                    spread_col,
                    spread,
                )

            if spread_pct is not None:
                ws.cell(
                    excel_row,
                    pct_col,
                    spread_pct / 100,
                )

        ws.cell(
            excel_row,
            1,
        ).number_format = "dd.mm.yyyy"

        ws.cell(
            excel_row,
            2,
        ).number_format = "hh:mm:ss"

        for col in (
            4, 5, 6,
            8, 9, 10,
            12, 13, 14,
        ):
            ws.cell(
                excel_row,
                col,
            ).number_format = "#,##0.0000"

        for col in (7, 11, 15):
            ws.cell(
                excel_row,
                col,
            ).number_format = "0.00%"

    if ws.max_row >= 2:

        _apply_table(
            ws,
            1,
            ws.max_row,
            len(headers),
            "GuncelKurlarTable",
        )

        for col_letter in ("G", "K", "O"):

            formula = (
                f'AND('
                f'{col_letter}2<>"",'
                f'{col_letter}2='
                f'MIN(${col_letter}$2:'
                f'${col_letter}${ws.max_row})'
                f')'
            )

            rule = FormulaRule(
                formula=[formula],
                fill=BEST_FILL,
                font=BEST_FONT,
            )

            ws.conditional_formatting.add(
                f"{col_letter}2:"
                f"{col_letter}{ws.max_row}",
                rule,
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
# GEÇMİŞ
# ============================================================

def _build_history_sheet(
    wb,
    history,
):

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

    _style_header(
        ws,
        1,
        1,
        len(headers),
    )

    sorted_history = sorted(
        history,
        key=_history_sort_key,
    )

    for excel_row, item in enumerate(
        sorted_history,
        start=2,
    ):

        run_dt = _parse_dt(
            item.get("run_at")
        )

        scraped_dt = _parse_dt(
            item.get("scraped_at")
        )

        code = (
            item.get("code")
            or ""
        ).strip()

        buy = _to_float(item.get("buy"))
        sell = _to_float(item.get("sell"))
        spread = _to_float(item.get("spread"))
        pct = _to_float(item.get("spread_pct"))

        site_spread = _to_float(
            item.get("site_spread")
        )

        site_pct = _to_float(
            item.get("site_spread_pct")
        )

        if (
            spread is None
            and buy is not None
            and sell is not None
        ):
            spread = sell - buy

        if (
            pct is None
            and spread is not None
            and buy not in (None, 0)
        ):
            pct = (spread / buy) * 100

        raw_status = item.get("status", "")

        ws.cell(
            excel_row,
            1,
            run_dt.date() if run_dt else "",
        )

        ws.cell(
            excel_row,
            2,
            (
                run_dt.time().replace(tzinfo=None)
                if run_dt
                else ""
            ),
        )

        provider = item.get("provider", "")

        provider_cell = ws.cell(
            excel_row,
            3,
            provider,
        )

        provider_cell.fill = _provider_fill(provider)
        provider_cell.font = Font(bold=True)

        product_fill = _product_fill(code)

        ws.cell(
            excel_row,
            4,
            item.get("product", ""),
        )

        for col in range(4, 9):
            ws.cell(
                excel_row,
                col,
            ).fill = product_fill

        ws.cell(
            excel_row,
            4,
        ).font = Font(bold=True)

        ws.cell(
            excel_row,
            5,
            buy if buy is not None else "",
        )

        ws.cell(
            excel_row,
            6,
            sell if sell is not None else "",
        )

        ws.cell(
            excel_row,
            7,
            spread if spread is not None else "",
        )

        ws.cell(
            excel_row,
            8,
            pct / 100 if pct is not None else "",
        )

        ws.cell(
            excel_row,
            9,
            _display_status(raw_status),
        )

        ws.cell(
            excel_row,
            10,
            item.get("source_url", ""),
        )

        ws.cell(
            excel_row,
            11,
            site_spread
            if site_spread is not None
            else "",
        )

        ws.cell(
            excel_row,
            12,
            site_pct / 100
            if site_pct is not None
            else "",
        )

        ws.cell(
            excel_row,
            13,
            item.get("note", ""),
        )

        ws.cell(
            excel_row,
            14,
            (
                scraped_dt.time().replace(
                    tzinfo=None
                )
                if scraped_dt
                else ""
            ),
        )

        # HAFTA SONU
        if (
            run_dt
            and run_dt.weekday() in (5, 6)
        ):
            date_cell = ws.cell(
                excel_row,
                1,
            )

            date_cell.fill = WEEKEND_FILL
            date_cell.font = Font(
                bold=True,
                color="9C6500",
            )

        ws.cell(
            excel_row,
            1,
        ).number_format = "dd.mm.yyyy"

        ws.cell(
            excel_row,
            2,
        ).number_format = "hh:mm:ss"

        ws.cell(
            excel_row,
            14,
        ).number_format = "hh:mm:ss"

        for col in (5, 6, 7, 11):
            ws.cell(
                excel_row,
                col,
            ).number_format = "#,##0.0000"

        for col in (8, 12):
            ws.cell(
                excel_row,
                col,
            ).number_format = "0.00%"

        source = ws.cell(
            excel_row,
            10,
        )

        if source.value:
            source.hyperlink = source.value
            source.style = "Hyperlink"

        status = ws.cell(
            excel_row,
            9,
        )

        status.fill = _status_fill(
            str(raw_status)
        )

        status.font = Font(bold=True)

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

    bank_trends = {}

    for row in history:

        provider = (
            row.get("provider")
            or ""
        ).strip()

        code = (
            row.get("code")
            or ""
        ).strip()

        run_at = row.get("run_at")
        run_dt = _parse_dt(run_at)

        if provider not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if not run_at or not run_dt:
            continue

        if row.get("status") == "ERROR":
            continue

        pct = _to_float(
            row.get("spread_pct")
        )

        if pct is None:
            buy = _to_float(row.get("buy"))
            sell = _to_float(row.get("sell"))

            if (
                buy not in (None, 0)
                and sell is not None
            ):
                pct = (
                    (sell - buy) / buy
                ) * 100

        if pct is None:
            continue

        bucket = bank_trends.setdefault(
            run_at,
            {
                "dt": run_dt,
                "banks": {},
            },
        )

        bucket["banks"].setdefault(
            provider,
            {},
        )[code] = pct / 100

    return sorted(
        bank_trends.values(),
        key=lambda x: x["dt"],
    )


# ============================================================
# AYLIK ORTALAMA
# ============================================================

def _monthly_averages(history):

    buckets = {}

    for row in history:

        provider = (
            row.get("provider")
            or ""
        ).strip()

        code = (
            row.get("code")
            or ""
        ).strip()

        dt = _parse_dt(
            row.get("run_at")
        )

        if provider not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if not dt:
            continue

        if row.get("status") == "ERROR":
            continue

        pct = _to_float(
            row.get("spread_pct")
        )

        if pct is None:
            buy = _to_float(row.get("buy"))
            sell = _to_float(row.get("sell"))

            if (
                buy not in (None, 0)
                and sell is not None
            ):
                pct = (
                    (sell - buy) / buy
                ) * 100

        if pct is None:
            continue

        key = (
            dt.year,
            dt.month,
            provider,
            code,
        )

        buckets.setdefault(
            key,
            [],
        ).append(
            pct / 100
        )

    result = {}

    for key, values in buckets.items():

        year, month, provider, code = key

        result.setdefault(
            (year, month),
            {},
        ).setdefault(
            provider,
            {},
        )[code] = (
            sum(values) / len(values)
        )

    return result


# ============================================================
# GRAFİK ORTAK STİLİ
# ============================================================

def _style_chart(
    chart,
    title,
    legend=False,
):

    chart.style = 10
    chart.title = title
    chart.height = 8.5
    chart.width = 21.0

    chart.y_axis.title = "Makas %"
    chart.x_axis.title = None

    try:
        chart.x_axis.axPos = "b"
        chart.x_axis.delete = False
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.tickLblSkip = 1
        chart.x_axis.tickMarkSkip = 1
        chart.x_axis.majorTickMark = "none"
        chart.x_axis.minorTickMark = "none"

        chart.y_axis.majorGridlines = None
        chart.y_axis.majorTickMark = "none"
        chart.y_axis.minorTickMark = "none"
        chart.y_axis.numFmt = "0.00%"

    except Exception:
        pass

    if legend:
        try:
            chart.legend.position = "b"
            chart.legend.overlay = False
        except Exception:
            pass
    else:
        chart.legend = None

    x_axis_text = RichText(
        bodyPr=RichTextProperties(
            rot=-2700000
        ),
        p=[
            Paragraph(
                pPr=ParagraphProperties(
                    defRPr=CharacterProperties(
                        sz=650
                    )
                ),
                endParaRPr=CharacterProperties(
                    sz=650
                ),
            )
        ],
    )

    try:
        chart.x_axis.txPr = x_axis_text
    except Exception:
        pass


# ============================================================
# ÖZET
# ============================================================

def _build_summary_sheet(
    wb,
    latest_run_at,
    latest_rows,
    history,
):

    ws = wb.create_sheet("OZET")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")

    ws["A1"] = (
        "Döviz ve Altın Kur Takip Özeti"
    )

    ws["A1"].fill = TITLE_FILL

    ws["A1"].font = Font(
        bold=True,
        size=16,
        color="1F4E78",
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    # ========================================================
    # TREND
    # ========================================================

    trend_runs = _build_bank_trends(
        history
    )

    category_values = []
    category_dates = []

    for run in trend_runs:

        dt = run["dt"]

        suffix = (
            " Cmt"
            if dt.weekday() == 5
            else
            " Paz"
            if dt.weekday() == 6
            else ""
        )

        category_values.append(
            dt.strftime("%d.%m")
            + suffix
        )

        category_dates.append(dt)

    # ========================================================
    # TOPLU GRAFİK HELPER
    # ========================================================

    helper_row = 500

    helper_columns = {}

    col = 2

    ws.cell(
        helper_row,
        1,
        "Tarih",
    )

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        helper_columns[code] = {}

        for bank in TARGET_BANKS:

            helper_columns[
                code
            ][bank] = col

            # Başlık yalnızca banka adı.
            # Böylece legend USD_Garanti değil,
            # Garanti BBVA olarak görünür.
            ws.cell(
                helper_row,
                col,
                bank,
            )

            col += 1

    for index, run in enumerate(
        trend_runs,
        start=helper_row + 1,
    ):

        label = category_values[
            index - helper_row - 1
        ]

        ws.cell(
            index,
            1,
            label,
        )

        for code in (
            "USD",
            "EUR",
            "XAU",
        ):

            for bank in TARGET_BANKS:

                value = (
                    run["banks"]
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:
                    ws.cell(
                        index,
                        helper_columns[
                            code
                        ][bank],
                        value,
                    ).number_format = (
                        "0.00%"
                    )

    helper_end = (
        helper_row
        + len(trend_runs)
    )

    # ========================================================
    # 1) ÜSTTE 3 TOPLU GÜNLÜK GRAFİK
    # ========================================================

    top_positions = [
        "A3",
        "J3",
        "T3",
    ]

    for product_index, code in enumerate(
        (
            "USD",
            "EUR",
            "XAU",
        )
    ):

        chart = LineChart()

        _style_chart(
            chart,
            (
                f"{PRODUCT_NAMES[code]} - "
                "5 Banka Makas %"
            ),
            legend=True,
        )

        cached = []

        for bank in TARGET_BANKS:

            data_col = (
                helper_columns[
                    code
                ][bank]
            )

            data = Reference(
                ws,
                min_col=data_col,
                max_col=data_col,
                min_row=helper_row,
                max_row=helper_end,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            values = [
                run["banks"]
                .get(bank, {})
                .get(code)
                for run in trend_runs
            ]

            cached.append(
                (bank, values)
            )

        cats = Reference(
            ws,
            min_col=1,
            min_row=helper_row + 1,
            max_row=helper_end,
        )

        chart.set_categories(cats)

        for index, series in enumerate(
            chart.series
        ):

            bank = TARGET_BANKS[index]

            try:
                series.graphicalProperties.line.solidFill = (
                    BANK_LINE_COLORS[bank]
                )

                series.graphicalProperties.line.width = (
                    20000
                )

                series.marker.symbol = "circle"
                series.marker.size = 5

                _apply_weekend_markers(
                    series,
                    category_dates,
                )

            except Exception:
                pass

        _cache_line_chart(
            chart,
            category_values,
            cached,
        )

        ws.add_chart(
            chart,
            top_positions[
                product_index
            ],
        )

    # ========================================================
    # 2) AYLIK ORTALAMA TABLOSU
    # ========================================================

    monthly = _monthly_averages(
        history
    )

    months = sorted(
        monthly.keys()
    )

    monthly_title_row = 26

    ws.merge_cells(
        start_row=monthly_title_row,
        start_column=1,
        end_row=monthly_title_row,
        end_column=5,
    )

    title = ws.cell(
        monthly_title_row,
        1,
        "AYLIK ORTALAMA MAKAS %",
    )

    title.fill = TITLE_FILL
    title.font = Font(
        bold=True,
        size=13,
        color="1F4E78",
    )

    title.alignment = Alignment(
        horizontal="center"
    )

    headers = [
        "Ay",
        "Banka",
        "Dolar Ort.",
        "Euro Ort.",
        "Gram Altın Ort.",
    ]

    for c, header in enumerate(
        headers,
        start=1,
    ):
        ws.cell(
            27,
            c,
            header,
        )

    _style_header(
        ws,
        27,
        1,
        5,
    )

    ws["C27"].fill = USD_HEADER_FILL
    ws["D27"].fill = EUR_HEADER_FILL
    ws["E27"].fill = XAU_HEADER_FILL

    table_row = 28

    for year, month in months:

        month_label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        for bank in TARGET_BANKS:

            ws.cell(
                table_row,
                1,
                month_label,
            )

            bank_cell = ws.cell(
                table_row,
                2,
                bank,
            )

            bank_cell.fill = _provider_fill(
                bank
            )

            bank_cell.font = Font(
                bold=True
            )

            for column, code in (
                (3, "USD"),
                (4, "EUR"),
                (5, "XAU"),
            ):

                cell = ws.cell(
                    table_row,
                    column,
                )

                cell.fill = _product_fill(
                    code
                )

                value = (
                    monthly
                    .get((year, month), {})
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:
                    cell.value = value
                    cell.number_format = (
                        "0.00%"
                    )

            for c in range(1, 6):
                ws.cell(
                    table_row,
                    c,
                ).border = THIN_BORDER

            table_row += 1

    # ========================================================
    # AYLIK HELPER
    # ========================================================

    monthly_helper = 600

    monthly_columns = {}

    ws.cell(
        monthly_helper,
        1,
        "Ay",
    )

    col = 2

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        monthly_columns[
            code
        ] = {}

        for bank in TARGET_BANKS:

            monthly_columns[
                code
            ][bank] = col

            ws.cell(
                monthly_helper,
                col,
                bank,
            )

            col += 1

    monthly_labels = []

    for index, month_key in enumerate(
        months,
        start=monthly_helper + 1,
    ):

        year, month = month_key

        label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        monthly_labels.append(label)

        ws.cell(
            index,
            1,
            label,
        )

        for code in (
            "USD",
            "EUR",
            "XAU",
        ):

            for bank in TARGET_BANKS:

                value = (
                    monthly
                    .get(month_key, {})
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    ws.cell(
                        index,
                        monthly_columns[
                            code
                        ][bank],
                        value,
                    ).number_format = (
                        "0.00%"
                    )

    monthly_end = (
        monthly_helper
        + len(months)
    )

    # ========================================================
    # 3) AYLIK 3 GRAFİK
    # ========================================================

    monthly_chart_positions = [
        "A42",
        "J42",
        "T42",
    ]

    if months:

        for product_index, code in enumerate(
            (
                "USD",
                "EUR",
                "XAU",
            )
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{PRODUCT_NAMES[code]} - "
                    "Aylık Ortalama"
                ),
                legend=True,
            )

            cached = []

            for bank in TARGET_BANKS:

                data = Reference(
                    ws,
                    min_col=monthly_columns[
                        code
                    ][bank],
                    max_col=monthly_columns[
                        code
                    ][bank],
                    min_row=monthly_helper,
                    max_row=monthly_end,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = [
                    monthly
                    .get(month_key, {})
                    .get(bank, {})
                    .get(code)
                    for month_key in months
                ]

                cached.append(
                    (
                        bank,
                        values,
                    )
                )

            cats = Reference(
                ws,
                min_col=1,
                min_row=monthly_helper + 1,
                max_row=monthly_end,
            )

            chart.set_categories(cats)

            for index, series in enumerate(
                chart.series
            ):

                bank = TARGET_BANKS[index]

                try:
                    series.graphicalProperties.line.solidFill = (
                        BANK_LINE_COLORS[
                            bank
                        ]
                    )

                    series.graphicalProperties.line.width = (
                        20000
                    )

                    series.marker.symbol = (
                        "circle"
                    )

                    series.marker.size = 6

                except Exception:
                    pass

            _cache_line_chart(
                chart,
                monthly_labels,
                cached,
            )

            ws.add_chart(
                chart,
                monthly_chart_positions[
                    product_index
                ],
            )

    # ========================================================
    # 4) MEVCUT 15 TEKİL GRAFİK
    # ========================================================

    # Üstte yeni iki grafik satırı olduğu için mevcut
    # banka grafiklerini blok halinde aşağı taşıyoruz.
    #
    # Grafiklerin kendi boyutları ve yatay düzeni AYNI.
    chart_columns = [
        "A",
        "J",
        "T",
    ]

    chart_row_starts = [
        65,
        88,
        111,
        134,
        157,
    ]

    # Her banka için kendi helper tablosu
    individual_helper_base = 700
    helper_gap = len(trend_runs) + 4

    for bank_index, bank in enumerate(
        TARGET_BANKS
    ):

        helper_start = (
            individual_helper_base
            + bank_index
            * helper_gap
        )

        ws.cell(
            helper_start,
            1,
            "Çekim Zamanı",
        )

        ws.cell(
            helper_start,
            2,
            "DOLAR",
        )

        ws.cell(
            helper_start,
            3,
            "EURO",
        )

        ws.cell(
            helper_start,
            4,
            "GRAM ALTIN",
        )

        cached_values = {
            "USD": [],
            "EUR": [],
            "XAU": [],
        }

        for offset, run in enumerate(
            trend_runs,
            start=1,
        ):

            row = helper_start + offset
            dt = run["dt"]

            suffix = (
                " Cmt"
                if dt.weekday() == 5
                else
                " Paz"
                if dt.weekday() == 6
                else ""
            )

            label = (
                dt.strftime("%d.%m")
                + suffix
                + " "
                + dt.strftime("%H:%M")
            )

            ws.cell(
                row,
                1,
                label,
            )

            for col, code in enumerate(
                (
                    "USD",
                    "EUR",
                    "XAU",
                ),
                start=2,
            ):

                value = (
                    run["banks"]
                    .get(bank, {})
                    .get(code)
                )

                cached_values[
                    code
                ].append(
                    value
                )

                if value is not None:
                    ws.cell(
                        row,
                        col,
                        value,
                    ).number_format = (
                        "0.00%"
                    )

        helper_end = (
            helper_start
            + len(trend_runs)
        )

        for product_index, code in enumerate(
            (
                "USD",
                "EUR",
                "XAU",
            )
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{bank} - "
                    f"{PRODUCT_NAMES[code]} "
                    "Makas %"
                ),
                legend=False,
            )

            bank_bg = (
                BANK_CHART_COLORS.get(
                    bank,
                    DEFAULT_BANK_CHART_COLOR,
                )
            )

            try:
                chart.graphical_properties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=bank_bg,
                    )
                )

                chart.plot_area.graphicalProperties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=bank_bg,
                    )
                )
            except Exception:
                pass

            data = Reference(
                ws,
                min_col=2 + product_index,
                max_col=2 + product_index,
                min_row=helper_start,
                max_row=helper_end,
            )

            cats = Reference(
                ws,
                min_col=1,
                min_row=helper_start + 1,
                max_row=helper_end,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            chart.set_categories(cats)

            try:
                series = chart.series[0]

                series.graphicalProperties.line.solidFill = (
                    PRODUCT_LINE_COLORS[
                        code
                    ]
                )

                series.graphicalProperties.line.width = (
                    20000
                )

                series.marker.symbol = (
                    "circle"
                    if code == "USD"
                    else
                    "square"
                    if code == "EUR"
                    else
                    "triangle"
                )

                series.marker.size = 6

                _apply_weekend_markers(
                    series,
                    category_dates,
                )

            except Exception:
                pass

            chart.dLbls = DataLabelList()
            chart.dLbls.showVal = True
            chart.dLbls.numFmt = "0.00%"
            chart.dLbls.dLblPos = "t"
            chart.dLbls.showLegendKey = False
            chart.dLbls.showCatName = False
            chart.dLbls.showSerName = False

            _cache_line_chart(
                chart,
                category_values,
                [
                    (
                        PRODUCT_NAMES[
                            code
                        ],
                        cached_values[
                            code
                        ],
                    )
                ],
            )

            values = [
                value
                for value
                in cached_values[
                    code
                ]
                if value is not None
            ]

            if values:
                minimum = min(values)
                maximum = max(values)

                padding = max(
                    (
                        maximum
                        - minimum
                    )
                    * 0.20,
                    maximum * 0.02,
                    0.00005,
                )

                chart.y_axis.scaling.min = max(
                    0,
                    minimum - padding,
                )

                chart.y_axis.scaling.max = (
                    maximum + padding
                )

            ws.add_chart(
                chart,
                (
                    f"{chart_columns[product_index]}"
                    f"{chart_row_starts[bank_index]}"
                ),
            )

    # ========================================================
    # 5) SON ÇEKİM BİLGİLERİ
    # ========================================================

    run_dt = _parse_dt(
        latest_run_at
    )

    providers = {
        row.get("provider")
        for row in latest_rows
        if row.get("provider")
    }

    summary_row = 183

    labels = [
        "Son Çekim Tarihi",
        "Son Çekim Saati",
        "Toplam Sağlayıcı",
        "Son Çekim Toplam Kayıt",
        "HATA",
        "KONTROL GEREKLİ",
    ]

    values = [
        run_dt.date()
        if run_dt
        else "",

        run_dt.time().replace(
            tzinfo=None
        )
        if run_dt
        else "",

        len(providers),

        len(latest_rows),

        sum(
            row.get("status")
            == "ERROR"
            for row in latest_rows
        ),

        sum(
            row.get("status")
            == "KONTROL"
            for row in latest_rows
        ),
    ]

    for offset, (
        label,
        value,
    ) in enumerate(
        zip(labels, values)
    ):

        row = summary_row + offset

        ws.cell(
            row,
            1,
            label,
        )

        ws.cell(
            row,
            1,
        ).font = Font(
            bold=True
        )

        ws.cell(
            row,
            1,
        ).fill = TITLE_FILL

        ws.cell(
            row,
            2,
            value,
        )

        ws.cell(
            row,
            1,
        ).border = THIN_BORDER

        ws.cell(
            row,
            2,
        ).border = THIN_BORDER

    ws.cell(
        summary_row,
        2,
    ).number_format = "dd.mm.yyyy"

    ws.cell(
        summary_row + 1,
        2,
    ).number_format = "hh:mm:ss"

    # ========================================================
    # 6) EN DÜŞÜK MAKAS - SADECE 5 BANKA
    # ========================================================

    best_header_row = 192

    best_headers = [
        "Ürün",
        "Banka Sayısı",
        "En Düşük Makas %",
        "Banka",
        "Alış",
        "Satış",
    ]

    for col, header in enumerate(
        best_headers,
        start=1,
    ):
        ws.cell(
            best_header_row,
            col,
            header,
        )

    _style_header(
        ws,
        best_header_row,
        1,
        6,
    )

    row_no = (
        best_header_row + 1
    )

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        valid = []

        for item in latest_rows:

            provider = (
                item.get("provider")
                or ""
            ).strip()

            if provider not in TARGET_BANKS:
                continue

            if item.get("code") != code:
                continue

            if item.get("status") == "ERROR":
                continue

            buy = _to_float(
                item.get("buy")
            )

            sell = _to_float(
                item.get("sell")
            )

            pct = _to_float(
                item.get("spread_pct")
            )

            if (
                pct is None
                and buy not in (None, 0)
                and sell is not None
            ):
                pct = (
                    (sell - buy) / buy
                ) * 100

            if (
                pct is not None
                and buy is not None
                and sell is not None
            ):
                valid.append(
                    (
                        pct,
                        provider,
                        buy,
                        sell,
                    )
                )

        best = (
            min(
                valid,
                key=lambda x: x[0],
            )
            if valid
            else None
        )

        product_cell = ws.cell(
            row_no,
            1,
            PRODUCT_NAMES[code],
        )

        product_cell.fill = (
            _product_fill(code)
        )

        product_cell.font = Font(
            bold=True
        )

        ws.cell(
            row_no,
            2,
            len(valid),
        )

        if best:

            pct, provider, buy, sell = best

            ws.cell(
                row_no,
                3,
                pct / 100,
            ).number_format = "0.00%"

            bank_cell = ws.cell(
                row_no,
                4,
                provider,
            )

            bank_cell.fill = (
                _provider_fill(provider)
            )

            bank_cell.font = Font(
                bold=True
            )

            ws.cell(
                row_no,
                5,
                buy,
            ).number_format = "#,##0.0000"

            ws.cell(
                row_no,
                6,
                sell,
            ).number_format = "#,##0.0000"

        for col in range(1, 7):
            ws.cell(
                row_no,
                col,
            ).border = THIN_BORDER

        row_no += 1

    # ========================================================
    # SÜTUN GENİŞLİKLERİ
    # ========================================================

    _set_widths(
        ws,
        {
            "A": 22,
            "B": 18,
            "C": 18,
            "D": 18,
            "E": 20,
            "F": 16,
            "G": 18,
            "H": 16,
            "I": 20,
            "J": 16,
            "K": 16,
            "L": 16,
            "M": 16,
            "N": 17,
            "O": 17,
            "P": 17,
            "Q": 17,
            "R": 17,
            "S": 17,
        },
    )


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

    # ========================================================
    # ÖNEMLİ AYRIM
    #
    # GUNCEL_KURLAR:
    # Gerçek EN SON run kullanılır.
    #
    # Böylece Hepsipay / Papara gibi son çekimde bulunan
    # sağlayıcılar kaybolmaz.
    #
    # GECMIS + GRAFİK + AYLIK:
    # Her günün yalnızca İLK run'ı kullanılır.
    # ========================================================

    latest_run_at, latest_rows = (
        _latest_run_rows(
            raw_history
        )
    )

    daily_history = (
        _daily_history(
            raw_history
        )
    )

    if not daily_history:
        raise RuntimeError(
            "Geçerli günlük geçmiş veri bulunamadı."
        )

    wb = Workbook()

    default_sheet = wb.active
    wb.remove(default_sheet)

    # 1) Güncel = gerçek son run
    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )

    # 2) Geçmiş = her gün ilk run
    _build_history_sheet(
        wb,
        daily_history,
    )

    # 3) Özet/grafikler = her gün ilk run
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
