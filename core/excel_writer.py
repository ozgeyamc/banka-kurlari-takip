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


STATUS_LABELS = {
    "OK": "DOĞRU",
    "KONTROL": "KONTROL GEREKLİ",
    "ERROR": "HATA",
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
    "Yapıkredi": "EACDCF",
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
    "Yapıkredi": "EDE4F7",
    "Ziraat Bankası": "FFF2CC",
    "İş Bankası": "DDEBF7",
}

DEFAULT_BANK_CHART_COLOR = "E8EDF3"


# ============================================================
# ÜRÜN RENKLERİ
#
# Bütün tek banka grafiklerinde:
#
# USD = MAVİ
# EUR = TURUNCU
# XAU = ALTIN
# ============================================================

PRODUCT_LINE_COLORS = {
    "USD": "4472C4",
    "EUR": "ED7D31",
    "XAU": "FFC000",
}


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


# ============================================================
# KARŞILAŞTIRILACAK 5 BANKA
# ============================================================

TARGET_BANKS = [
    "Garanti BBVA",
    "Akbank",
    "Yapıkredi",
    "Ziraat Bankası",
    "İş Bankası",
]


# Toplu grafiklerde bankaları birbirinden ayıran renkler.
COMPARISON_BANK_COLORS = {
    "Garanti BBVA": "70AD47",
    "Akbank": "C00000",
    "Yapıkredi": "7030A0",
    "Ziraat Bankası": "FFC000",
    "İş Bankası": "4472C4",
}


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def _provider_fill(provider: str | None):
    color = PROVIDER_COLORS.get(
        (provider or "").strip(),
        DEFAULT_PROVIDER_COLOR,
    )

    return PatternFill(
        "solid",
        fgColor=color,
    )


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

    with path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        return list(csv.DictReader(handle))


def _style_header(
    ws,
    row: int,
    start_col: int,
    end_col: int,
) -> None:

    for col in range(
        start_col,
        end_col + 1,
    ):
        cell = ws.cell(
            row=row,
            column=col,
        )

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        cell.border = THIN_BORDER


def _apply_table(
    ws,
    start_row: int,
    end_row: int,
    end_col: int,
    name: str,
) -> None:

    if end_row <= start_row:
        return

    from openpyxl.utils import get_column_letter

    ref = (
        f"A{start_row}:"
        f"{get_column_letter(end_col)}"
        f"{end_row}"
    )

    table = Table(
        displayName=name,
        ref=ref,
    )

    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    ws.add_table(table)


def _set_widths(
    ws,
    widths: dict[str, float],
) -> None:

    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def _status_fill(status: str):

    if status == "ERROR":
        return ERROR_FILL

    if status == "KONTROL":
        return CONTROL_FILL

    return OK_FILL


def _display_status(
    status: str | None,
) -> str:

    raw = (status or "").strip()

    return STATUS_LABELS.get(
        raw,
        raw,
    )


# ============================================================
# GÜNLÜK TEK ÇALIŞTIRMA
#
# ÖNEMLİ:
#
# Aynı gün birden fazla çalışma varsa
# günün İLK çalıştırması tutulur.
#
# 08:00 -> KALIR
# 10:00 -> ATILIR
# 15:00 -> ATILIR
#
# CSV fiziksel olarak silinmez.
# Excel tarafında filtrelenir.
# ============================================================

def _daily_history(
    history: list[dict],
) -> list[dict]:

    first_run_by_date: dict = {}

    for row in history:

        run_dt = _parse_dt(
            row.get("run_at")
        )

        if not run_dt:
            continue

        date_key = run_dt.date()

        current = first_run_by_date.get(
            date_key
        )

        if (
            current is None
            or run_dt < current
        ):
            first_run_by_date[
                date_key
            ] = run_dt

    selected_datetimes = set(
        first_run_by_date.values()
    )

    result = []

    for row in history:

        run_dt = _parse_dt(
            row.get("run_at")
        )

        if run_dt in selected_datetimes:
            result.append(row)

    return result


def _latest_run_rows(
    history: list[dict],
) -> tuple[str | None, list[dict]]:

    valid = [
        row
        for row in history
        if _parse_dt(
            row.get("run_at")
        )
    ]

    if not valid:
        return None, []

    latest = max(
        valid,
        key=lambda row: _parse_dt(
            row.get("run_at")
        ),
    )

    latest_run_at = latest[
        "run_at"
    ]

    return (
        latest_run_at,
        [
            row
            for row in history
            if row.get("run_at")
            == latest_run_at
        ],
    )


def _provider_map(
    rows: Iterable[dict],
) -> dict[str, dict[str, dict]]:

    result: dict[
        str,
        dict[str, dict],
    ] = {}

    for row in rows:

        provider = (
            row.get(
                "provider",
                "",
            )
            .strip()
        )

        code = (
            row.get(
                "code",
                "",
            )
            .strip()
        )

        if not provider or not code:
            continue

        result.setdefault(
            provider,
            {},
        )[code] = row

    return result


def _history_sort_key(
    row: dict,
):

    provider = (
        row.get("provider")
        or ""
    ).strip().casefold()

    code = (
        row.get("code")
        or ""
    ).strip()

    run_dt = (
        _parse_dt(
            row.get("run_at")
        )
        or datetime.min
    )

    return (
        provider,
        PRODUCT_ORDER.get(
            code,
            99,
        ),
        run_dt,
    )


# ============================================================
# CHART CACHE
# ============================================================

def _cache_line_chart(
    chart,
    categories: list[str],
    series_values: list[
        tuple[
            str,
            list[float | None],
        ]
    ],
) -> None:

    for series, (
        title,
        values,
    ) in zip(
        chart.series,
        series_values,
    ):

        numeric_points = [
            NumVal(
                idx=index,
                v=float(value),
            )
            for index, value
            in enumerate(values)
            if value is not None
        ]

        if (
            series.val is not None
            and series.val.numRef
            is not None
        ):

            series.val.numRef.numCache = (
                NumData(
                    formatCode="0.000%",
                    ptCount=len(values),
                    pt=numeric_points,
                )
            )

        if series.cat is not None:

            category_formula = None

            if (
                series.cat.numRef
                is not None
            ):
                category_formula = (
                    series.cat.numRef.f
                )

            elif (
                series.cat.strRef
                is not None
            ):
                category_formula = (
                    series.cat.strRef.f
                )

            series.cat.numRef = None

            series.cat.strRef = StrRef(
                f=category_formula,
                strCache=StrData(
                    ptCount=len(
                        categories
                    ),
                    pt=[
                        StrVal(
                            idx=index,
                            v=str(value),
                        )
                        for index, value
                        in enumerate(
                            categories
                        )
                    ],
                ),
            )

        if (
            series.tx is not None
            and series.tx.strRef
            is not None
        ):

            series.tx.strRef.strCache = (
                StrData(
                    ptCount=1,
                    pt=[
                        StrVal(
                            idx=0,
                            v=title,
                        )
                    ],
                )
            )


# ============================================================
# GÜNCEL KURLAR
# ============================================================

def _build_current_sheet(
    wb: Workbook,
    latest_run_at: str | None,
    latest_rows: list[dict],
) -> None:

    ws = wb.create_sheet(
        "GUNCEL_KURLAR"
    )

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

    provider_map = _provider_map(
        latest_rows
    )

    run_dt = _parse_dt(
        latest_run_at
    )

    for excel_row, provider in enumerate(
        sorted(
            provider_map,
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

        provider_cell.fill = (
            _provider_fill(
                provider
            )
        )

        provider_cell.font = Font(
            bold=True
        )

        layout = {
            "USD": (
                4,
                5,
                6,
                7,
            ),
            "EUR": (
                8,
                9,
                10,
                11,
            ),
            "XAU": (
                12,
                13,
                14,
                15,
            ),
        }

        for code, (
            buy_col,
            sell_col,
            spread_col,
            pct_col,
        ) in layout.items():

            item = row_map.get(
                code
            )

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
                item.get(
                    "spread_pct"
                )
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
                and buy
                not in (
                    None,
                    0,
                )
            ):
                spread_pct = (
                    spread
                    / buy
                ) * 100.0

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
                    spread_pct
                    / 100.0,
                )

        ws.cell(
            excel_row,
            1,
        ).number_format = (
            "dd.mm.yyyy"
        )

        ws.cell(
            excel_row,
            2,
        ).number_format = (
            "hh:mm:ss"
        )

        for col in (
            4,
            5,
            6,
            8,
            9,
            10,
            12,
            13,
            14,
        ):

            ws.cell(
                excel_row,
                col,
            ).number_format = (
                "#,##0.0000"
            )

        for col in (
            7,
            11,
            15,
        ):

            ws.cell(
                excel_row,
                col,
            ).number_format = (
                "0.00%"
            )

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
                formula=[
                    formula
                ],
                fill=BEST_FILL,
                font=BEST_FONT,
            )

            ws.conditional_formatting.add(
                f"{col_letter}2:"
                f"{col_letter}"
                f"{ws.max_row}",
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
    wb: Workbook,
    history: list[dict],
) -> None:

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
            item.get(
                "scraped_at"
            )
        )

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
            item.get(
                "spread_pct"
            )
        )

        site_spread = _to_float(
            item.get(
                "site_spread"
            )
        )

        site_spread_pct = _to_float(
            item.get(
                "site_spread_pct"
            )
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
            and buy
            not in (
                None,
                0,
            )
        ):
            spread_pct = (
                spread / buy
            ) * 100.0

        raw_status = item.get(
            "status",
            "",
        )

        ws.cell(
            excel_row,
            1,
            (
                run_dt.date()
                if run_dt
                else ""
            ),
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

        provider_name = (
            item.get(
                "provider",
                "",
            )
        )

        provider_cell = ws.cell(
            excel_row,
            3,
            provider_name,
        )

        provider_cell.fill = (
            _provider_fill(
                provider_name
            )
        )

        provider_cell.font = Font(
            bold=True
        )

        ws.cell(
            excel_row,
            4,
            item.get(
                "product",
                "",
            ),
        )

        ws.cell(
            excel_row,
            5,
            (
                buy
                if buy is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            6,
            (
                sell
                if sell is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            7,
            (
                spread
                if spread is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            8,
            (
                spread_pct / 100.0
                if spread_pct
                is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            9,
            _display_status(
                raw_status
            ),
        )

        ws.cell(
            excel_row,
            10,
            item.get(
                "source_url",
                "",
            ),
        )

        ws.cell(
            excel_row,
            11,
            (
                site_spread
                if site_spread
                is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            12,
            (
                site_spread_pct
                / 100.0
                if site_spread_pct
                is not None
                else ""
            ),
        )

        ws.cell(
            excel_row,
            13,
            item.get(
                "note",
                "",
            ),
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

        ws.cell(
            excel_row,
            1,
        ).number_format = (
            "dd.mm.yyyy"
        )

        ws.cell(
            excel_row,
            2,
        ).number_format = (
            "hh:mm:ss"
        )

        ws.cell(
            excel_row,
            14,
        ).number_format = (
            "hh:mm:ss"
        )

        for col in (
            5,
            6,
            7,
            11,
        ):

            ws.cell(
                excel_row,
                col,
            ).number_format = (
                "#,##0.0000"
            )

        for col in (
            8,
            12,
        ):

            ws.cell(
                excel_row,
                col,
            ).number_format = (
                "0.00%"
            )

        source_cell = ws.cell(
            excel_row,
            10,
        )

        if source_cell.value:
            source_cell.hyperlink = (
                source_cell.value
            )
            source_cell.style = (
                "Hyperlink"
            )

        status_cell = ws.cell(
            excel_row,
            9,
        )

        status_cell.fill = (
            _status_fill(
                str(
                    raw_status
                )
            )
        )

        status_cell.font = Font(
            bold=True
        )

        status_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
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

    for row in ws.iter_rows(
        min_row=2,
        min_col=10,
        max_col=10,
    ):

        row[0].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

    for row in ws.iter_rows(
        min_row=2,
        min_col=13,
        max_col=13,
    ):

        row[0].alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )


# ============================================================
# ÖZET
# ============================================================

def _build_summary_sheet(
    wb: Workbook,
    latest_run_at: str | None,
    latest_rows: list[dict],
    history: list[dict],
) -> None:

    ws = wb.create_sheet(
        "OZET"
    )

    ws.sheet_view.showGridLines = False

    # ========================================================
    # BAŞLIK
    # ========================================================

    ws.merge_cells(
        "A1:AD1"
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

    ws.row_dimensions[1].height = 28

    run_dt = _parse_dt(
        latest_run_at
    )

    # ========================================================
    # TREND VERİSİ
    # ========================================================

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

        run_dt_item = _parse_dt(
            row.get("run_at")
        )

        if provider not in TARGET_BANKS:
            continue

        if code not in (
            "USD",
            "EUR",
            "XAU",
        ):
            continue

        if not run_dt_item:
            continue

        if (
            row.get("status")
            == "ERROR"
        ):
            continue

        spread_pct = _to_float(
            row.get(
                "spread_pct"
            )
        )

        if spread_pct is None:

            buy = _to_float(
                row.get("buy")
            )

            sell = _to_float(
                row.get("sell")
            )

            if (
                buy not in (
                    None,
                    0,
                )
                and sell is not None
            ):

                spread_pct = (
                    (
                        sell - buy
                    )
                    / buy
                ) * 100.0

        if spread_pct is None:
            continue

        date_key = (
            run_dt_item.date()
        )

        run_bucket = (
            bank_trends.setdefault(
                date_key,
                {
                    "dt": run_dt_item,
                    "banks": {},
                },
            )
        )

        bank_bucket = (
            run_bucket[
                "banks"
            ].setdefault(
                provider,
                {},
            )
        )

        bank_bucket[
            code
        ] = (
            spread_pct
            / 100.0
        )

    trend_runs = sorted(
        bank_trends.values(),
        key=lambda item: item[
            "dt"
        ],
    )

    # ========================================================
    # YARDIMCI VERİ TABLOSU
    # ========================================================

    helper_base_row = 220

    helper_headers = [
        "Tarih"
    ]

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        for bank in TARGET_BANKS:

            helper_headers.append(
                f"{code}_{bank}"
            )

    for col, header in enumerate(
        helper_headers,
        start=1,
    ):

        ws.cell(
            helper_base_row,
            col,
            header,
        )

    category_values = []

    for helper_row, run in enumerate(
        trend_runs,
        start=helper_base_row + 1,
    ):

        label = run[
            "dt"
        ].strftime(
            "%d.%m.%Y"
        )

        category_values.append(
            label
        )

        ws.cell(
            helper_row,
            1,
            label,
        )

        col = 2

        for code in (
            "USD",
            "EUR",
            "XAU",
        ):

            for bank in TARGET_BANKS:

                value = (
                    run["banks"]
                    .get(
                        bank,
                        {},
                    )
                    .get(code)
                )

                cell = ws.cell(
                    helper_row,
                    col,
                )

                if value is not None:

                    cell.value = (
                        value
                    )

                    cell.number_format = (
                        "0.00%"
                    )

                col += 1

    helper_end_row = (
        helper_base_row
        + len(
            trend_runs
        )
    )

    # ========================================================
    # ORTAK CHART STİLİ
    # ========================================================

    def style_chart(
        chart,
        title,
        show_legend=False,
    ):

        chart.style = 10

        chart.title = title

        chart.height = 8.5

        chart.width = 21.0

        chart.y_axis.title = (
            "Makas %"
        )

        chart.x_axis.title = None

        try:

            chart.x_axis.axPos = (
                "b"
            )

            chart.x_axis.delete = (
                False
            )

            chart.x_axis.tickLblPos = (
                "low"
            )

            chart.x_axis.tickLblSkip = (
                1
            )

            chart.x_axis.tickMarkSkip = (
                1
            )

            chart.x_axis.majorTickMark = (
                "none"
            )

            chart.x_axis.minorTickMark = (
                "none"
            )

        except Exception:
            pass

        try:

            chart.y_axis.majorGridlines = (
                None
            )

            chart.y_axis.majorTickMark = (
                "none"
            )

            chart.y_axis.minorTickMark = (
                "none"
            )

        except Exception:
            pass

        if not show_legend:
            chart.legend = None

        try:
            chart.y_axis.numFmt = (
                "0.00%"
            )
        except Exception:
            pass

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

        y_axis_text = RichText(
            bodyPr=RichTextProperties(),
            p=[
                Paragraph(
                    pPr=ParagraphProperties(
                        defRPr=CharacterProperties(
                            sz=800
                        )
                    ),
                    endParaRPr=CharacterProperties(
                        sz=800
                    ),
                )
            ],
        )

        try:

            chart.x_axis.txPr = (
                x_axis_text
            )

            chart.y_axis.txPr = (
                y_axis_text
            )

        except Exception:
            pass

    # ========================================================
    # ÜSTTEKİ 3 TOPLU GRAFİK
    # ========================================================

    top_chart_positions = [
        "A3",
        "J3",
        "T3",
    ]

    product_start_columns = {
        "USD": 2,
        "EUR": 7,
        "XAU": 12,
    }

    for product_index, code in enumerate(
        (
            "USD",
            "EUR",
            "XAU",
        )
    ):

        chart = LineChart()

        style_chart(
            chart,
            (
                f"{PRODUCT_NAMES[code]} - "
                "5 Banka Makas % Karşılaştırması"
            ),
            show_legend=True,
        )

        start_col = (
            product_start_columns[
                code
            ]
        )

        cached_series = []

        for bank_index, bank in enumerate(
            TARGET_BANKS
        ):

            data_col = (
                start_col
                + bank_index
            )

            data = Reference(
                ws,
                min_col=data_col,
                max_col=data_col,
                min_row=helper_base_row,
                max_row=helper_end_row,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            values = [
                run["banks"]
                .get(
                    bank,
                    {},
                )
                .get(code)
                for run
                in trend_runs
            ]

            cached_series.append(
                (
                    bank,
                    values,
                )
            )

        cats = Reference(
            ws,
            min_col=1,
            min_row=helper_base_row + 1,
            max_row=helper_end_row,
        )

        chart.set_categories(
            cats
        )

        # --------------------------------------------
        # TOPLU GRAFİKTE HER BANKA TEK RENK
        # --------------------------------------------

        for series_index, series in enumerate(
            chart.series
        ):

            bank = TARGET_BANKS[
                series_index
            ]

            color = (
                COMPARISON_BANK_COLORS[
                    bank
                ]
            )

            try:

                series.graphicalProperties.line.solidFill = (
                    color
                )

                series.graphicalProperties.line.width = (
                    25000
                )

                series.marker.symbol = (
                    "circle"
                )

                series.marker.size = 5

                series.marker.graphicalProperties.solidFill = (
                    color
                )

                series.marker.graphicalProperties.line.solidFill = (
                    color
                )

            except Exception:
                pass

        _cache_line_chart(
            chart,
            category_values,
            cached_series,
        )

        all_values = []

        for _, values in cached_series:

            all_values.extend(
                [
                    value
                    for value in values
                    if value is not None
                ]
            )

        if all_values:

            minimum = min(
                all_values
            )

            maximum = max(
                all_values
            )

            padding = max(
                (
                    maximum
                    - minimum
                ) * 0.15,
                maximum * 0.02,
                0.00005,
            )

            chart.y_axis.scaling.min = (
                max(
                    0,
                    minimum
                    - padding,
                )
            )

            chart.y_axis.scaling.max = (
                maximum
                + padding
            )

        ws.add_chart(
            chart,
            top_chart_positions[
                product_index
            ],
        )

    # ========================================================
    # 15 TEK BANKA GRAFİĞİ
    # ========================================================

    chart_columns = [
        "A",
        "J",
        "T",
    ]

    chart_row_starts = [
        27,
        50,
        73,
        96,
        119,
    ]

    for bank_index, bank in enumerate(
        TARGET_BANKS
    ):

        for product_index, code in enumerate(
            (
                "USD",
                "EUR",
                "XAU",
            )
        ):

            chart = LineChart()

            style_chart(
                chart,
                (
                    f"{bank} - "
                    f"{PRODUCT_NAMES[code]} "
                    "Makas %"
                ),
            )

            data_col = (
                product_start_columns[
                    code
                ]
                + TARGET_BANKS.index(
                    bank
                )
            )

            data = Reference(
                ws,
                min_col=data_col,
                max_col=data_col,
                min_row=helper_base_row,
                max_row=helper_end_row,
            )

            cats = Reference(
                ws,
                min_col=1,
                min_row=helper_base_row + 1,
                max_row=helper_end_row,
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
                .get(
                    bank,
                    {},
                )
                .get(code)
                for run
                in trend_runs
            ]

            # --------------------------------------------
            # ÜRÜNE GÖRE TEK RENK
            # --------------------------------------------

            try:

                series = (
                    chart.series[0]
                )

                color = (
                    PRODUCT_LINE_COLORS[
                        code
                    ]
                )

                series.graphicalProperties.line.solidFill = (
                    color
                )

                series.graphicalProperties.line.width = (
                    25000
                )

                series.marker.symbol = (
                    "circle"
                )

                series.marker.size = 6

                series.marker.graphicalProperties.solidFill = (
                    color
                )

                series.marker.graphicalProperties.line.solidFill = (
                    color
                )

            except Exception:
                pass

            # --------------------------------------------
            # BANKA ARKA PLANI
            # --------------------------------------------

            background_color = (
                BANK_CHART_COLORS.get(
                    bank,
                    DEFAULT_BANK_CHART_COLOR,
                )
            )

            try:

                chart.graphical_properties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=background_color,
                    )
                )

            except Exception:
                pass

            try:

                chart.plot_area.graphicalProperties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=background_color,
                    )
                )

            except Exception:
                pass

            # --------------------------------------------
            # NOKTA ÜZERİNDE DEĞER
            # --------------------------------------------

            chart.dLbls = (
                DataLabelList()
            )

            chart.dLbls.showVal = True

            chart.dLbls.numFmt = (
                "0.00%"
            )

            chart.dLbls.dLblPos = "t"

            chart.dLbls.showLegendKey = (
                False
            )

            chart.dLbls.showCatName = (
                False
            )

            chart.dLbls.showSerName = (
                False
            )

            try:
                chart.dLbls.showLeaderLines = (
                    False
                )
            except Exception:
                pass

            data_label_text = RichText(
                bodyPr=RichTextProperties(),
                p=[
                    Paragraph(
                        pPr=ParagraphProperties(
                            defRPr=CharacterProperties(
                                sz=700
                            )
                        ),
                        endParaRPr=CharacterProperties(
                            sz=700
                        ),
                    )
                ],
            )

            try:
                chart.dLbls.txPr = (
                    data_label_text
                )
            except Exception:
                pass

            _cache_line_chart(
                chart,
                category_values,
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
                    ) * 0.20,
                    maximum * 0.02,
                    0.00005,
                )

                chart.y_axis.scaling.min = (
                    max(
                        0,
                        minimum
                        - padding,
                    )
                )

                chart.y_axis.scaling.max = (
                    maximum
                    + padding
                )

            position = (
                f"{chart_columns[product_index]}"
                f"{chart_row_starts[bank_index]}"
            )

            ws.add_chart(
                chart,
                position,
            )

    # ========================================================
    # ÇEKİM ÖZETİ
    # ========================================================

    summary_start = 145

    providers = {
        row.get("provider")
        for row in latest_rows
        if row.get("provider")
    }

    error_count = sum(
        row.get("status")
        == "ERROR"
        for row in latest_rows
    )

    control_count = sum(
        row.get("status")
        == "KONTROL"
        for row in latest_rows
    )

    summary_labels = [
        "Son Çekim Tarihi",
        "Son Çekim Saati",
        "Toplam Sağlayıcı",
        "Son Çekim Toplam Kayıt",
        "HATA",
        "KONTROL GEREKLİ",
    ]

    summary_values = [
        (
            run_dt.date()
            if run_dt
            else ""
        ),
        (
            run_dt.time().replace(
                tzinfo=None
            )
            if run_dt
            else ""
        ),
        len(providers),
        len(latest_rows),
        error_count,
        control_count,
    ]

    for offset, (
        label,
        value,
    ) in enumerate(
        zip(
            summary_labels,
            summary_values,
        )
    ):

        row_no = (
            summary_start
            + offset
        )

        label_cell = ws.cell(
            row_no,
            1,
            label,
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

        value_cell = ws.cell(
            row_no,
            2,
            value,
        )

        value_cell.border = (
            THIN_BORDER
        )

        value_cell.alignment = Alignment(
            horizontal="center"
        )

    ws.cell(
        summary_start,
        2,
    ).number_format = (
        "dd.mm.yyyy"
    )

    ws.cell(
        summary_start + 1,
        2,
    ).number_format = (
        "hh:mm:ss"
    )

    # ========================================================
    # SADECE 5 BANKA ARASINDA EN DÜŞÜK MAKAS
    # ========================================================

    best_table_row = (
        summary_start + 8
    )

    ws.merge_cells(
        start_row=best_table_row - 1,
        start_column=1,
        end_row=best_table_row - 1,
        end_column=6,
    )

    best_title = ws.cell(
        best_table_row - 1,
        1,
    )

    best_title.value = (
        "5 BANKA ARASINDA EN DÜŞÜK MAKAS"
    )

    best_title.font = Font(
        bold=True,
        size=12,
        color="1F4E78",
    )

    best_title.fill = TITLE_FILL

    best_title.alignment = Alignment(
        horizontal="center"
    )

    headers = [
        "Ürün",
        "Banka Sayısı",
        "En Düşük Makas %",
        "Banka",
        "Alış",
        "Satış",
    ]

    for col, header in enumerate(
        headers,
        start=1,
    ):

        ws.cell(
            best_table_row,
            col,
            header,
        )

    _style_header(
        ws,
        best_table_row,
        1,
        6,
    )

    current_row = (
        best_table_row + 1
    )

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        product_rows = [
            row
            for row in latest_rows
            if (
                row.get("code")
                == code
                and (
                    row.get(
                        "provider"
                    )
                    or ""
                ).strip()
                in TARGET_BANKS
            )
        ]

        valid_rows = []

        for item in product_rows:

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
                item.get(
                    "spread_pct"
                )
            )

            if (
                pct is None
                and buy
                not in (
                    None,
                    0,
                )
                and sell is not None
            ):

                pct = (
                    (
                        sell - buy
                    )
                    / buy
                ) * 100.0

            if (
                buy
                not in (
                    None,
                    0,
                )
                and sell is not None
                and pct is not None
            ):

                valid_rows.append(
                    (
                        pct,
                        item,
                        buy,
                        sell,
                    )
                )

        best = (
            min(
                valid_rows,
                key=lambda item: item[
                    0
                ],
            )
            if valid_rows
            else None
        )

        ws.cell(
            current_row,
            1,
            PRODUCT_NAMES[
                code
            ],
        )

        bank_count = len(
            {
                (
                    row.get(
                        "provider"
                    )
                    or ""
                ).strip()
                for row
                in product_rows
            }
        )

        ws.cell(
            current_row,
            2,
            bank_count,
        )

        if best:

            pct, item, buy, sell = (
                best
            )

            ws.cell(
                current_row,
                3,
                pct / 100.0,
            )

            provider_name = (
                item.get(
                    "provider"
                )
                or ""
            ).strip()

            bank_cell = ws.cell(
                current_row,
                4,
                provider_name,
            )

            bank_cell.fill = (
                _provider_fill(
                    provider_name
                )
            )

            bank_cell.font = Font(
                bold=True
            )

            ws.cell(
                current_row,
                5,
                buy,
            )

            ws.cell(
                current_row,
                6,
                sell,
            )

            ws.cell(
                current_row,
                3,
            ).number_format = (
                "0.00%"
            )

            ws.cell(
                current_row,
                5,
            ).number_format = (
                "#,##0.0000"
            )

            ws.cell(
                current_row,
                6,
            ).number_format = (
                "#,##0.0000"
            )

        for col in range(
            1,
            7,
        ):

            ws.cell(
                current_row,
                col,
            ).border = (
                THIN_BORDER
            )

        current_row += 1

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
            "T": 17,
            "U": 17,
            "V": 17,
            "W": 17,
            "X": 17,
            "Y": 17,
            "Z": 17,
            "AA": 17,
            "AB": 17,
            "AC": 17,
            "AD": 17,
        },
    )


# ============================================================
# ANA EXCEL OLUŞTURMA FONKSİYONU
# ============================================================

def build_excel(
    history_path: str | Path,
    output_path: str | Path,
) -> None:

    history = read_history(
        history_path
    )

    if not history:

        raise RuntimeError(
            "Excel oluşturmak için "
            "geçmiş veri bulunamadı."
        )

    # ========================================================
    # KRİTİK:
    #
    # Aynı gün birden fazla çalıştırma varsa
    # yalnızca günün İLK çalıştırmasını kullan.
    # ========================================================

    daily_history = _daily_history(
        history
    )

    if not daily_history:

        raise RuntimeError(
            "Geçerli günlük geçmiş "
            "verisi bulunamadı."
        )

    latest_run_at, latest_rows = (
        _latest_run_rows(
            daily_history
        )
    )

    wb = Workbook()

    default_sheet = wb.active

    wb.remove(
        default_sheet
    )

    # ========================================================
    # 1) GÜNCEL
    # ========================================================

    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )

    # ========================================================
    # 2) GEÇMİŞ
    #
    # Buraya ham history değil,
    # günlük filtrelenmiş history gönderiliyor.
    # ========================================================

    _build_history_sheet(
        wb,
        daily_history,
    )

    # ========================================================
    # 3) ÖZET
    # ========================================================

    _build_summary_sheet(
        wb,
        latest_run_at,
        latest_rows,
        daily_history,
    )

    # Excel açıldığında ilk sayfa
    # GUNCEL_KURLAR olsun.
    wb.active = 0

    output_path = Path(
        output_path
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Her seferinde aynı Excel dosyası güncellenir.
    wb.save(
        output_path
    )
