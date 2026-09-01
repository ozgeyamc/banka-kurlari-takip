from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
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
from openpyxl.utils import get_column_letter


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
    "USD": "4472C4",   # mavi
    "EUR": "ED7D31",   # turuncu
    "XAU": "FFC000",   # altın
}


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

    # Hepsipay ve Papara özellikle korunuyor.
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

    # Yapıkredi artık Akbank'tan belirgin farklı.
    "Yapıkredi": "D9C2E9",

    "Ziraat Bankası": "DFF1E3",
    "Ziraat Dinamik": "DBCAEC",
    "Ziraat Katılım": "F1EDDF",
}

DEFAULT_PROVIDER_COLOR = "E8EDF3"


# ============================================================
# BANKA GRAFİK ARKA PLANLARI
# ============================================================

BANK_CHART_COLORS = {
    "Garanti BBVA": "E2F0D9",
    "Akbank": "FDE2E2",
    "Yapıkredi": "E4D7F5",
    "Ziraat Bankası": "FFF2CC",
    "İş Bankası": "DDEBF7",
}

DEFAULT_BANK_CHART_COLOR = "E8EDF3"


# ============================================================
# 5 BANKA KARŞILAŞTIRMA RENKLERİ
# ============================================================

BANK_LINE_COLORS = {
    "Garanti BBVA": "70AD47",    # yeşil
    "Akbank": "C00000",          # kırmızı
    "Yapıkredi": "7030A0",       # mor
    "Ziraat Bankası": "FFC000",  # sarı
    "İş Bankası": "4472C4",      # mavi
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
# HELPER ALANI
#
# Grafiklerin kullandığı teknik verileri AA sütunundan itibaren
# yazıyoruz ve daha sonra bu sütunları gizliyoruz.
# Böylece OZET sayfasının en altında anlamsız teknik tablolar
# görünmüyor.
# ============================================================

HELPER_START_COL = 27  # AA


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
# TARİH ETİKETİ
#
# Normal gün:
# 28.08
#
# Cumartesi:
# 29.08 Cmt
#
# Pazar:
# 30.08 Paz
#
# Çizgi veya noktaların rengini hafta sonunda DEĞİŞTİRMİYORUZ.
# ============================================================

def _date_label(dt: datetime, include_time=False):

    if dt.weekday() == 5:
        suffix = " Cmt"

    elif dt.weekday() == 6:
        suffix = " Paz"

    else:
        suffix = ""

    label = dt.strftime("%d.%m") + suffix

    if include_time:
        label += " " + dt.strftime("%H:%M")

    return label


# ============================================================
# AYNI GÜNDE İLK ÇEKİM
# ============================================================

def _daily_history(history: list[dict]) -> list[dict]:

    runs_by_day = {}

    for row in history:

        raw = row.get("run_at")
        dt = _parse_dt(raw)

        if not dt or not raw:
            continue

        day = dt.date()

        current = runs_by_day.get(day)

        if current is None:
            runs_by_day[day] = raw
            continue

        current_dt = _parse_dt(current)

        if current_dt is None or dt < current_dt:
            runs_by_day[day] = raw

    selected_runs = set(
        runs_by_day.values()
    )

    return [
        row
        for row in history
        if row.get("run_at") in selected_runs
    ]


# ============================================================
# GERÇEK SON RUN
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

    latest_rows = [
        row
        for row in history
        if row.get("run_at") == latest_run_at
    ]

    return latest_run_at, latest_rows


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
        run_dt,
        provider,
        PRODUCT_ORDER.get(code, 99),
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
# ÇİZGİ + MARKER AYNI RENK
# ============================================================

def _set_series_color(
    series,
    color,
    marker="circle",
    marker_size=6,
):
    """
    Çizgi ve marker aynı renk olur.

    Hafta sonu için marker rengi değiştirilmez.
    """

    try:
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 22000

        series.marker.symbol = marker
        series.marker.size = marker_size

        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color

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

    # Ürün başlık renkleri
    for col in range(4, 8):
        ws.cell(1, col).fill = USD_HEADER_FILL

    for col in range(8, 12):
        ws.cell(1, col).fill = EUR_HEADER_FILL

    for col in range(12, 16):
        ws.cell(1, col).fill = XAU_HEADER_FILL

    provider_map = _provider_map(
        latest_rows
    )

    run_dt = _parse_dt(
        latest_run_at
    )

    for excel_row, provider in enumerate(
        sorted(
            provider_map.keys(),
            key=str.casefold,
        ),
        start=2,
    ):

        row_map = provider_map[
            provider
        ]

        ws.cell(
            excel_row,
            1,
            run_dt.date()
            if run_dt
            else "",
        )

        ws.cell(
            excel_row,
            2,
            (
                run_dt.time().replace(
                    tzinfo=None
                )
                if run_dt
                else ""
            ),
        )

        provider_cell = ws.cell(
            excel_row,
            3,
            provider,
        )

        provider_cell.fill = _provider_fill(
            provider
        )

        provider_cell.font = Font(
            bold=True
        )

        layout = {
            "USD": (4, 5, 6, 7),
            "EUR": (8, 9, 10, 11),
            "XAU": (12, 13, 14, 15),
        }

        for code, cols in layout.items():

            buy_col, sell_col, spread_col, pct_col = cols

            fill = _product_fill(
                code
            )

            for col in cols:
                ws.cell(
                    excel_row,
                    col,
                ).fill = fill

            item = row_map.get(
                code
            )

            # Kaynakta bu ürün yoksa hücreler boş kalır.
            # Hepsipay/Papara yalnızca XAU veriyorsa
            # USD/EUR uydurulmaz.
            if not item:
                continue

            buy = _to_float(
                item.get("buy")
            )

            sell = _to_float(
                item.get("sell")
            )

            spread = _to_float(
                item.get("spread")
            )

            spread_pct = _to_float(
                item.get("spread_pct")
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

        for col in (
            7,
            11,
            15,
        ):
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

        for col_letter in (
            "G",
            "K",
            "O",
        ):

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

    ws = wb.create_sheet(
        "GECMIS"
    )

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

        buy = _to_float(
            item.get("buy")
        )

        sell = _to_float(
            item.get("sell")
        )

        spread = _to_float(
            item.get("spread")
        )

        pct = _to_float(
            item.get("spread_pct")
        )

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
            pct = (
                spread / buy
            ) * 100

        raw_status = (
            item.get("status", "")
        )

        ws.cell(
            excel_row,
            1,
            run_dt.date()
            if run_dt
            else "",
        )

        ws.cell(
            excel_row,
            2,
            (
                run_dt.time().replace(
                    tzinfo=None
                )
                if run_dt
                else ""
            ),
        )

        provider = (
            item.get("provider", "")
        )

        provider_cell = ws.cell(
            excel_row,
            3,
            provider,
        )

        provider_cell.fill = (
            _provider_fill(provider)
        )

        provider_cell.font = Font(
            bold=True
        )

        product_fill = (
            _product_fill(code)
        )

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
        ).font = Font(
            bold=True
        )

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
            pct / 100
            if pct is not None
            else "",
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
            (
                site_spread
                if site_spread is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            12,
            (
                site_pct / 100
                if site_pct is not None
                else ""
            ),
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

        # ----------------------------------------------------
        # HAFTA SONU
        # Sadece tarih hücresi renklendirilir.
        # Grafik çizgi/nokta rengine dokunulmaz.
        # ----------------------------------------------------

        if (
            run_dt
            and run_dt.weekday() in (5, 6)
        ):

            date_cell = ws.cell(
                excel_row,
                1,
            )

            date_cell.fill = (
                WEEKEND_FILL
            )

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

        for col in (
            5,
            6,
            7,
            11,
        ):
            ws.cell(
                excel_row,
                col,
            ).number_format = "#,##0.0000"

        for col in (
            8,
            12,
        ):
            ws.cell(
                excel_row,
                col,
            ).number_format = "0.00%"

        source = ws.cell(
            excel_row,
            10,
        )

        if source.value:
            source.hyperlink = (
                source.value
            )
            source.style = "Hyperlink"

        status = ws.cell(
            excel_row,
            9,
        )

        status.fill = (
            _status_fill(
                str(raw_status)
            )
        )

        status.font = Font(
            bold=True
        )

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

def _build_bank_trends(
    history
):

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

        run_at = (
            row.get("run_at")
        )

        run_dt = _parse_dt(
            run_at
        )

        if provider not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if not run_at or not run_dt:
            continue

        if (
            row.get("status")
            == "ERROR"
        ):
            continue

        pct = _to_float(
            row.get("spread_pct")
        )

        if pct is None:

            buy = _to_float(
                row.get("buy")
            )

            sell = _to_float(
                row.get("sell")
            )

            if (
                buy not in (None, 0)
                and sell is not None
            ):
                pct = (
                    (sell - buy)
                    / buy
                ) * 100

        if pct is None:
            continue

        bucket = (
            bank_trends.setdefault(
                run_at,
                {
                    "dt": run_dt,
                    "banks": {},
                },
            )
        )

        bucket[
            "banks"
        ].setdefault(
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

def _monthly_averages(
    history
):

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

        if (
            row.get("status")
            == "ERROR"
        ):
            continue

        pct = _to_float(
            row.get("spread_pct")
        )

        if pct is None:

            buy = _to_float(
                row.get("buy")
            )

            sell = _to_float(
                row.get("sell")
            )

            if (
                buy not in (None, 0)
                and sell is not None
            ):
                pct = (
                    (sell - buy)
                    / buy
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

        (
            year,
            month,
            provider,
            code,
        ) = key

        result.setdefault(
            (year, month),
            {},
        ).setdefault(
            provider,
            {},
        )[code] = (
            sum(values)
            / len(values)
        )

    return result


# ============================================================
# GRAFİK ORTAK STİL
# ============================================================

def _style_chart(
    chart,
    title,
    legend=False,
):

    chart.style = 10
    chart.title = title

    chart.height = 10.0
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

    try:

        chart.x_axis.txPr = RichText(
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

    ws = wb.create_sheet(
        "OZET"
    )

    ws.sheet_view.showGridLines = False

    # ========================================================
    # BAŞLIK
    # ========================================================

    ws.merge_cells(
        "A1:F1"
    )

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
    # TREND VERİLERİ
    # ========================================================

    trend_runs = (
        _build_bank_trends(
            history
        )
    )

    category_values = [
        _date_label(
            run["dt"],
            include_time=False,
        )
        for run in trend_runs
    ]

    # ========================================================
    # GÜNLÜK TOPLU GRAFİK HELPER
    #
    # AA sütunundan başlıyor.
    # ========================================================

    daily_helper_row = 2

    daily_date_col = (
        HELPER_START_COL
    )

    ws.cell(
        daily_helper_row,
        daily_date_col,
        "Tarih",
    )

    daily_columns = {}

    current_col = (
        daily_date_col + 1
    )

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        daily_columns[
            code
        ] = {}

        for bank in TARGET_BANKS:

            daily_columns[
                code
            ][bank] = (
                current_col
            )

            ws.cell(
                daily_helper_row,
                current_col,
                bank,
            )

            current_col += 1

    for row_offset, run in enumerate(
        trend_runs,
        start=1,
    ):

        row = (
            daily_helper_row
            + row_offset
        )

        ws.cell(
            row,
            daily_date_col,
            _date_label(
                run["dt"],
                include_time=False,
            ),
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

                    cell = ws.cell(
                        row,
                        daily_columns[
                            code
                        ][bank],
                        value,
                    )

                    cell.number_format = (
                        "0.00%"
                    )

    daily_helper_end = (
        daily_helper_row
        + len(trend_runs)
    )

    # ========================================================
    # 1) ÜST BÖLÜM
    # 5 BANKA TOPLU GÜNLÜK GRAFİKLER
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
                daily_columns[
                    code
                ][bank]
            )

            data = Reference(
                ws,
                min_col=data_col,
                max_col=data_col,
                min_row=daily_helper_row,
                max_row=daily_helper_end,
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
                (
                    bank,
                    values,
                )
            )

        cats = Reference(
            ws,
            min_col=daily_date_col,
            min_row=daily_helper_row + 1,
            max_row=daily_helper_end,
        )

        chart.set_categories(
            cats
        )

        # ----------------------------------------------------
        # HER BANKA ÇİZGİSİ + NOKTASI AYNI RENK
        #
        # Hafta sonu için renk DEĞİŞMİYOR.
        # Sadece tarih Cmt / Paz olarak yazıyor.
        # ----------------------------------------------------

        for index, series in enumerate(
            chart.series
        ):

            bank = TARGET_BANKS[
                index
            ]

            _set_series_color(
                series,
                BANK_LINE_COLORS[
                    bank
                ],
                marker="circle",
                marker_size=6,
            )

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

    monthly = (
        _monthly_averages(
            history
        )
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

    monthly_title = ws.cell(
        monthly_title_row,
        1,
        "AYLIK ORTALAMA MAKAS %",
    )

    monthly_title.fill = (
        TITLE_FILL
    )

    monthly_title.font = Font(
        bold=True,
        size=13,
        color="1F4E78",
    )

    monthly_title.alignment = Alignment(
        horizontal="center"
    )

    monthly_headers = [
        "Ay",
        "Banka",
        "Dolar Ort.",
        "Euro Ort.",
        "Gram Altın Ort.",
    ]

    for col, header in enumerate(
        monthly_headers,
        start=1,
    ):
        ws.cell(
            27,
            col,
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

    monthly_table_row = 28

    for year, month in months:

        month_label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        for bank in TARGET_BANKS:

            ws.cell(
                monthly_table_row,
                1,
                month_label,
            )

            bank_cell = ws.cell(
                monthly_table_row,
                2,
                bank,
            )

            bank_cell.fill = (
                _provider_fill(
                    bank
                )
            )

            bank_cell.font = Font(
                bold=True
            )

            for col, code in (
                (3, "USD"),
                (4, "EUR"),
                (5, "XAU"),
            ):

                cell = ws.cell(
                    monthly_table_row,
                    col,
                )

                cell.fill = (
                    _product_fill(code)
                )

                value = (
                    monthly
                    .get(
                        (year, month),
                        {},
                    )
                    .get(
                        bank,
                        {},
                    )
                    .get(code)
                )

                if value is not None:
                    cell.value = value
                    cell.number_format = (
                        "0.00%"
                    )

            for col in range(
                1,
                6,
            ):
                ws.cell(
                    monthly_table_row,
                    col,
                ).border = (
                    THIN_BORDER
                )

            monthly_table_row += 1

    # ========================================================
    # AYLIK GRAFİK HELPER
    # ========================================================

    monthly_helper_row = 100

    monthly_date_col = (
        HELPER_START_COL
    )

    ws.cell(
        monthly_helper_row,
        monthly_date_col,
        "Ay",
    )

    monthly_columns = {}

    current_col = (
        monthly_date_col + 1
    )

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
            ][bank] = (
                current_col
            )

            ws.cell(
                monthly_helper_row,
                current_col,
                bank,
            )

            current_col += 1

    monthly_labels = []

    for row_offset, month_key in enumerate(
        months,
        start=1,
    ):

        year, month = month_key

        label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        monthly_labels.append(
            label
        )

        row = (
            monthly_helper_row
            + row_offset
        )

        ws.cell(
            row,
            monthly_date_col,
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
                    .get(
                        month_key,
                        {},
                    )
                    .get(
                        bank,
                        {},
                    )
                    .get(code)
                )

                if value is not None:

                    cell = ws.cell(
                        row,
                        monthly_columns[
                            code
                        ][bank],
                        value,
                    )

                    cell.number_format = (
                        "0.00%"
                    )

    monthly_helper_end = (
        monthly_helper_row
        + len(months)
    )

    # ========================================================
    # 3) AYLIK ORTALAMA GRAFİKLER
    # ========================================================

    monthly_positions = [
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
                    "Aylık Ortalama Makas %"
                ),
                legend=True,
            )

            cached = []

            for bank in TARGET_BANKS:

                data_col = (
                    monthly_columns[
                        code
                    ][bank]
                )

                data = Reference(
                    ws,
                    min_col=data_col,
                    max_col=data_col,
                    min_row=monthly_helper_row,
                    max_row=monthly_helper_end,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = [
                    monthly
                    .get(
                        month_key,
                        {},
                    )
                    .get(
                        bank,
                        {},
                    )
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
                min_col=monthly_date_col,
                min_row=monthly_helper_row + 1,
                max_row=monthly_helper_end,
            )

            chart.set_categories(
                cats
            )

            # ------------------------------------------------
            # AYLIK ORTALAMA:
            #
            # Garanti:
            # çizgi yeşil + nokta yeşil
            #
            # Akbank:
            # çizgi kırmızı + nokta kırmızı
            #
            # Yapıkredi:
            # çizgi mor + nokta mor
            #
            # Ziraat:
            # çizgi sarı + nokta sarı
            #
            # İş Bankası:
            # çizgi mavi + nokta mavi
            # ------------------------------------------------

            for index, series in enumerate(
                chart.series
            ):

                bank = TARGET_BANKS[
                    index
                ]

                _set_series_color(
                    series,
                    BANK_LINE_COLORS[
                        bank
                    ],
                    marker="circle",
                    marker_size=7,
                )

            _cache_line_chart(
                chart,
                monthly_labels,
                cached,
            )

            ws.add_chart(
                chart,
                monthly_positions[
                    product_index
                ],
            )

    # ========================================================
    # 4) TEKİL BANKA GRAFİKLERİ
    # ========================================================

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

    individual_helper_row = 200

    individual_date_col = (
        HELPER_START_COL
    )

    # Her banka için 4 sütun:
    # Tarih | USD | EUR | XAU
    bank_helper_columns = {}

    current_col = (
        individual_date_col
    )

    for bank in TARGET_BANKS:

        bank_helper_columns[
            bank
        ] = {
            "DATE": current_col,
            "USD": current_col + 1,
            "EUR": current_col + 2,
            "XAU": current_col + 3,
        }

        ws.cell(
            individual_helper_row,
            current_col,
            f"{bank} Tarih",
        )

        ws.cell(
            individual_helper_row,
            current_col + 1,
            "DOLAR",
        )

        ws.cell(
            individual_helper_row,
            current_col + 2,
            "EURO",
        )

        ws.cell(
            individual_helper_row,
            current_col + 3,
            "GRAM ALTIN",
        )

        current_col += 4

    for row_offset, run in enumerate(
        trend_runs,
        start=1,
    ):

        row = (
            individual_helper_row
            + row_offset
        )

        for bank in TARGET_BANKS:

            cols = (
                bank_helper_columns[
                    bank
                ]
            )

            ws.cell(
                row,
                cols["DATE"],
                _date_label(
                    run["dt"],
                    include_time=True,
                ),
            )

            for code in (
                "USD",
                "EUR",
                "XAU",
            ):

                value = (
                    run["banks"]
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    cell = ws.cell(
                        row,
                        cols[code],
                        value,
                    )

                    cell.number_format = (
                        "0.00%"
                    )

    individual_helper_end = (
        individual_helper_row
        + len(trend_runs)
    )

    for bank_index, bank in enumerate(
        TARGET_BANKS
    ):

        cols = (
            bank_helper_columns[
                bank
            ]
        )

        individual_categories = [
            _date_label(
                run["dt"],
                include_time=True,
            )
            for run in trend_runs
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
                    f"{bank} - "
                    f"{PRODUCT_NAMES[code]} "
                    "Makas %"
                ),
                legend=False,
            )

            # Bankaya göre pastel arka plan
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
                min_col=cols[code],
                max_col=cols[code],
                min_row=individual_helper_row,
                max_row=individual_helper_end,
            )

            cats = Reference(
                ws,
                min_col=cols["DATE"],
                min_row=individual_helper_row + 1,
                max_row=individual_helper_end,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            chart.set_categories(
                cats
            )

            values = [
                run["banks"]
                .get(bank, {})
                .get(code)
                for run in trend_runs
            ]

            # ------------------------------------------------
            # TEKİL GRAFİKLER
            #
            # Dolar:
            # çizgi mavi + marker mavi
            #
            # Euro:
            # çizgi turuncu + marker turuncu
            #
            # Gram Altın:
            # çizgi altın + marker altın
            #
            # HAFTA SONUNDA RENK DEĞİŞMEZ.
            # ------------------------------------------------

            if chart.series:

                marker = (
                    "circle"
                    if code == "USD"
                    else
                    "square"
                    if code == "EUR"
                    else
                    "triangle"
                )

                _set_series_color(
                    chart.series[0],
                    PRODUCT_LINE_COLORS[
                        code
                    ],
                    marker=marker,
                    marker_size=6,
                )

            chart.dLbls = (
                DataLabelList()
            )

            chart.dLbls.showVal = True
            chart.dLbls.numFmt = "0.00%"
            chart.dLbls.dLblPos = "t"

            chart.dLbls.showLegendKey = False
            chart.dLbls.showCatName = False
            chart.dLbls.showSerName = False

            _cache_line_chart(
                chart,
                individual_categories,
                [
                    (
                        PRODUCT_NAMES[
                            code
                        ],
                        values,
                    )
                ],
            )

            valid_values = [
                value
                for value in values
                if value is not None
            ]

            if valid_values:

                minimum = min(
                    valid_values
                )

                maximum = max(
                    valid_values
                )

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
    # 5) ÖZET BİLGİLER
    #
    # Eski teknik görünen alan yerine daha anlaşılır bölüm.
    # ========================================================

    run_dt = _parse_dt(
        latest_run_at
    )

    providers = {
        (
            row.get("provider")
            or ""
        ).strip()
        for row in latest_rows
        if (
            row.get("provider")
            or ""
        ).strip()
    }

    summary_title_row = 183

    ws.merge_cells(
        start_row=summary_title_row,
        start_column=1,
        end_row=summary_title_row,
        end_column=4,
    )

    cell = ws.cell(
        summary_title_row,
        1,
        "SON VERİ ÇEKİMİ ÖZETİ",
    )

    cell.fill = TITLE_FILL

    cell.font = Font(
        bold=True,
        size=12,
        color="1F4E78",
    )

    cell.alignment = Alignment(
        horizontal="center"
    )

    summary_rows = [
        (
            "Veri tarihi",
            (
                run_dt.date()
                if run_dt
                else ""
            ),
        ),
        (
            "Veri çekim saati",
            (
                run_dt.time().replace(
                    tzinfo=None
                )
                if run_dt
                else ""
            ),
        ),
        (
            "Toplam sağlayıcı",
            len(providers),
        ),
        (
            "Bu çekimde alınan ürün kaydı",
            len(latest_rows),
        ),
        (
            "Hatalı kayıt",
            sum(
                (
                    row.get("status")
                    == "ERROR"
                )
                for row in latest_rows
            ),
        ),
        (
            "Kontrol edilmesi gereken kayıt",
            sum(
                (
                    row.get("status")
                    == "KONTROL"
                )
                for row in latest_rows
            ),
        ),
    ]

    row = (
        summary_title_row + 1
    )

    for label, value in summary_rows:

        label_cell = ws.cell(
            row,
            1,
            label,
        )

        value_cell = ws.cell(
            row,
            2,
            value,
        )

        label_cell.font = Font(
            bold=True
        )

        label_cell.fill = (
            PatternFill(
                "solid",
                fgColor="EAF2F8",
            )
        )

        label_cell.border = (
            THIN_BORDER
        )

        value_cell.border = (
            THIN_BORDER
        )

        row += 1

    ws.cell(
        summary_title_row + 1,
        2,
    ).number_format = "dd.mm.yyyy"

    ws.cell(
        summary_title_row + 2,
        2,
    ).number_format = "hh:mm:ss"

    # ========================================================
    # 6) 5 BANKADA EN DÜŞÜK MAKAS
    # ========================================================

    best_title_row = 191

    ws.merge_cells(
        start_row=best_title_row,
        start_column=1,
        end_row=best_title_row,
        end_column=6,
    )

    best_title = ws.cell(
        best_title_row,
        1,
        "5 BANKA ARASINDA EN DÜŞÜK GÜNCEL MAKAS",
    )

    best_title.fill = TITLE_FILL

    best_title.font = Font(
        bold=True,
        size=12,
        color="1F4E78",
    )

    best_title.alignment = Alignment(
        horizontal="center"
    )

    best_header_row = (
        best_title_row + 1
    )

    best_headers = [
        "Ürün",
        "Karşılaştırılan Banka",
        "En Düşük Makas %",
        "En Avantajlı Banka",
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

    best_row = (
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

            if (
                item.get("code")
                != code
            ):
                continue

            if (
                item.get("status")
                == "ERROR"
            ):
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
                    (sell - buy)
                    / buy
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
            best_row,
            1,
            PRODUCT_NAMES[
                code
            ],
        )

        product_cell.fill = (
            _product_fill(code)
        )

        product_cell.font = Font(
            bold=True
        )

        ws.cell(
            best_row,
            2,
            len(valid),
        )

        if best:

            (
                pct,
                provider,
                buy,
                sell,
            ) = best

            ws.cell(
                best_row,
                3,
                pct / 100,
            ).number_format = (
                "0.00%"
            )

            bank_cell = ws.cell(
                best_row,
                4,
                provider,
            )

            bank_cell.fill = (
                _provider_fill(
                    provider
                )
            )

            bank_cell.font = Font(
                bold=True
            )

            ws.cell(
                best_row,
                5,
                buy,
            ).number_format = (
                "#,##0.0000"
            )

            ws.cell(
                best_row,
                6,
                sell,
            ).number_format = (
                "#,##0.0000"
            )

        for col in range(
            1,
            7,
        ):
            ws.cell(
                best_row,
                col,
            ).border = (
                THIN_BORDER
            )

        best_row += 1

    # ========================================================
    # HELPER SÜTUNLARINI GİZLE
    #
    # AA ve sonrasındaki teknik grafik verileri artık kullanıcıya
    # görünmeyecek.
    # ========================================================

    last_helper_col = max(
        current_col,
        HELPER_START_COL + 20,
    )

    for col in range(
        HELPER_START_COL,
        last_helper_col + 1,
    ):
        ws.column_dimensions[
            get_column_letter(col)
        ].hidden = True

    # ========================================================
    # SÜTUN GENİŞLİKLERİ
    # ========================================================

    _set_widths(
        ws,
        {
            "A": 28,
            "B": 20,
            "C": 18,
            "D": 22,
            "E": 18,
            "F": 18,
            "G": 16,
            "H": 16,
            "I": 16,
            "J": 16,
            "K": 16,
            "L": 16,
            "M": 16,
            "N": 16,
            "O": 16,
            "P": 16,
            "Q": 16,
            "R": 16,
            "S": 16,
            "T": 16,
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

    # --------------------------------------------------------
    # GUNCEL_KURLAR
    #
    # Gerçek EN SON run kullanılır.
    #
    # Bu nedenle Hepsipay / Papara son çekimde varsa
    # güncel sayfada mutlaka yer alır.
    # --------------------------------------------------------

    (
        latest_run_at,
        latest_rows,
    ) = _latest_run_rows(
        raw_history
    )

    # --------------------------------------------------------
    # GECMIS + OZET + GRAFİK + AYLIK
    #
    # Her günün yalnızca İLK çekimi kullanılır.
    # --------------------------------------------------------

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
    wb.remove(
        default_sheet
    )

    # 1. Güncel
    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )

    # 2. Geçmiş
    _build_history_sheet(
        wb,
        daily_history,
    )

    # 3. Özet
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
