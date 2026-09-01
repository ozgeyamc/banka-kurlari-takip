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

# Hafta sonu işaretlemesi
WEEKEND_FILL = PatternFill("solid", fgColor="FFF2CC")

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
    "USD": "4472C4",  # mavi
    "EUR": "ED7D31",  # turuncu
    "XAU": "FFC000",  # altın
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
# ÜRÜN / BANKA AYARLARI
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

# Aylık karşılaştırma grafiklerinde banka renkleri
COMPARISON_BANK_COLORS = {
    "Garanti BBVA": "70AD47",
    "Akbank": "C00000",
    "Yapıkredi": "7030A0",
    "Ziraat Bankası": "FFC000",
    "İş Bankası": "4472C4",
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
        f"{get_column_letter(end_col)}{end_row}"
    )

    table = Table(
        displayName=name,
        ref=ref,
    )

    # Ürün hücrelerindeki kendi renklerimiz görünsün diye
    # satır çizgilerini kapatıyoruz.
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
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
# AYNI GÜN BİRDEN FAZLA ÇALIŞMA VARSA İLKİNİ AL
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

        day = run_dt.date()

        current = first_run_by_date.get(day)

        if (
            current is None
            or run_dt < current
        ):
            first_run_by_date[day] = run_dt

    selected_runs = set(
        first_run_by_date.values()
    )

    result = []

    for row in history:

        run_dt = _parse_dt(
            row.get("run_at")
        )

        if run_dt in selected_runs:
            result.append(row)

    return result


def _latest_run_rows(
    history: list[dict],
) -> tuple[str | None, list[dict]]:

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
            if row.get("run_at")
            == latest_run_at
        ],
    )


def _provider_map(
    rows: Iterable[dict],
) -> dict[str, dict[str, dict]]:

    result: dict[
        str,
        dict[str, dict]
    ] = {}

    for row in rows:

        provider = (
            row.get(
                "provider",
                "",
            ).strip()
        )

        code = (
            row.get(
                "code",
                "",
            ).strip()
        )

        if (
            not provider
            or not code
        ):
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

    # Ürün başlık renkleri
    for col in range(4, 8):
        ws.cell(
            1,
            col,
        ).fill = USD_HEADER_FILL

    for col in range(8, 12):
        ws.cell(
            1,
            col,
        ).fill = EUR_HEADER_FILL

    for col in range(12, 16):
        ws.cell(
            1,
            col,
        ).fill = XAU_HEADER_FILL

    provider_map = (
        _provider_map(
            latest_rows
        )
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

        row_map = (
            provider_map[
                provider
            ]
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
                run_dt.time()
                .replace(
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

            fill = _product_fill(
                code
            )

            # Ürün grubunun tamamını renklendir
            for col in (
                buy_col,
                sell_col,
                spread_col,
                pct_col,
            ):
                ws.cell(
                    excel_row,
                    col,
                ).fill = fill

            item = row_map.get(code)

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
                spread = (
                    sell - buy
                )

            if (
                spread_pct is None
                and spread is not None
                and buy not in (
                    None,
                    0,
                )
            ):
                spread_pct = (
                    spread / buy
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
                    spread_pct / 100.0,
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
            1,
        ).alignment = Alignment(
            horizontal="center"
        )

        ws.cell(
            excel_row,
            2,
        ).alignment = Alignment(
            horizontal="center"
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
                formula=[formula],
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

        spread_pct = _to_float(
            item.get("spread_pct")
        )

        site_spread = _to_float(
            item.get("site_spread")
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
            and buy not in (
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
                run_dt.time()
                .replace(
                    tzinfo=None
                )
                if run_dt
                else ""
            ),
        )

        provider_name = item.get(
            "provider",
            "",
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

        product_fill = (
            _product_fill(
                code
            )
        )

        product_cell = ws.cell(
            excel_row,
            4,
            item.get(
                "product",
                "",
            ),
        )

        product_cell.fill = (
            product_fill
        )

        product_cell.font = Font(
            bold=True
        )

        # Ürün + Alış + Satış + Makas + Makas %
        for col in range(
            4,
            9,
        ):
            ws.cell(
                excel_row,
                col,
            ).fill = product_fill

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
                site_spread_pct / 100.0
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
                scraped_dt.time()
                .replace(
                    tzinfo=None
                )
                if scraped_dt
                else ""
            ),
        )

        # Hafta sonu tarih hücresini belirginleştir.
        if (
            run_dt
            and run_dt.weekday()
            in (
                5,
                6,
            )
        ):

            ws.cell(
                excel_row,
                1,
            ).fill = WEEKEND_FILL

            ws.cell(
                excel_row,
                1,
            ).font = Font(
                bold=True
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

        ws.cell(
            excel_row,
            1,
        ).alignment = Alignment(
            horizontal="center"
        )

        ws.cell(
            excel_row,
            2,
        ).alignment = Alignment(
            horizontal="center"
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
                str(raw_status)
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
# AYLIK ORTALAMALAR
# ============================================================

def _build_monthly_averages(
    history: list[dict],
) -> dict:

    buckets: dict = {}

    for row in history:

        provider = (
            row.get("provider")
            or ""
        ).strip()

        code = (
            row.get("code")
            or ""
        ).strip()

        run_dt = _parse_dt(
            row.get("run_at")
        )

        if provider not in TARGET_BANKS:
            continue

        if code not in {
            "USD",
            "EUR",
            "XAU",
        }:
            continue

        if not run_dt:
            continue

        if (
            row.get("status")
            == "ERROR"
        ):
            continue

        spread_pct = _to_float(
            row.get("spread_pct")
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
                    (sell - buy)
                    / buy
                ) * 100.0

        if spread_pct is None:
            continue

        key = (
            run_dt.year,
            run_dt.month,
            provider,
            code,
        )

        buckets.setdefault(
            key,
            [],
        ).append(
            spread_pct / 100.0
        )

    result: dict = {}

    for (
        year,
        month,
        provider,
        code,
    ), values in buckets.items():

        if not values:
            continue

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

    # --------------------------------------------------------
    # BU KISIM SENİN MEVCUT DÜZENİNLE AYNI
    # --------------------------------------------------------

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

    ws.row_dimensions[1].height = 26

    run_dt = _parse_dt(
        latest_run_at
    )

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

    labels = [
        (
            "A121",
            "Son Çekim Tarihi",
        ),
        (
            "A122",
            "Son Çekim Saati",
        ),
        (
            "A123",
            "Toplam Sağlayıcı",
        ),
        (
            "A124",
            "Son Çekim Toplam Kayıt",
        ),
        (
            "A125",
            "HATA",
        ),
        (
            "A126",
            "KONTROL GEREKLİ",
        ),
    ]

    for coord, label in labels:

        ws[coord] = label
        ws[coord].font = Font(
            bold=True
        )

        ws[coord].fill = PatternFill(
            "solid",
            fgColor="EAF2F8",
        )

        ws[coord].border = (
            THIN_BORDER
        )

    ws["B121"] = (
        run_dt.date()
        if run_dt
        else ""
    )

    ws["B122"] = (
        run_dt.time()
        .replace(
            tzinfo=None
        )
        if run_dt
        else ""
    )

    ws["B123"] = len(
        providers
    )

    ws["B124"] = len(
        latest_rows
    )

    ws["B125"] = (
        error_count
    )

    ws["B126"] = (
        control_count
    )

    ws["B121"].number_format = (
        "dd.mm.yyyy"
    )

    ws["B122"].number_format = (
        "hh:mm:ss"
    )

    for row in range(
        121,
        127,
    ):

        ws[
            f"B{row}"
        ].border = THIN_BORDER

        ws[
            f"B{row}"
        ].alignment = Alignment(
            horizontal="center"
        )

    headers = [
        "Ürün",
        "Sağlayıcı Sayısı",
        "En Düşük Makas %",
        "Sağlayıcı",
        "Alış",
        "Satış",
    ]

    for col, header in enumerate(
        headers,
        1,
    ):

        ws.cell(
            row=129,
            column=col,
            value=header,
        )

    _style_header(
        ws,
        129,
        1,
        len(headers),
    )

    row_no = 130

    # SADECE 5 BANKA
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

        for row in product_rows:

            buy = _to_float(
                row.get("buy")
            )

            sell = _to_float(
                row.get("sell")
            )

            pct = _to_float(
                row.get(
                    "spread_pct"
                )
            )

            if (
                row.get("status")
                != "ERROR"
                and buy
                and sell
                and pct is not None
            ):
                valid_rows.append(
                    (
                        pct,
                        row,
                        buy,
                        sell,
                    )
                )

        best = (
            min(
                valid_rows,
                key=lambda item:
                item[0],
            )
            if valid_rows
            else None
        )

        product_cell = ws.cell(
            row=row_no,
            column=1,
            value=PRODUCT_NAMES[
                code
            ],
        )

        product_cell.fill = (
            _product_fill(
                code
            )
        )

        product_cell.font = Font(
            bold=True
        )

        ws.cell(
            row=row_no,
            column=2,
            value=len(
                product_rows
            ),
        )

        if best:

            pct, row, buy, sell = (
                best
            )

            ws.cell(
                row=row_no,
                column=3,
                value=pct / 100.0,
            )

            bank_cell = ws.cell(
                row=row_no,
                column=4,
                value=row.get(
                    "provider"
                ),
            )

            bank_cell.fill = (
                _provider_fill(
                    row.get(
                        "provider"
                    )
                )
            )

            bank_cell.font = Font(
                bold=True
            )

            ws.cell(
                row=row_no,
                column=5,
                value=buy,
            )

            ws.cell(
                row=row_no,
                column=6,
                value=sell,
            )

            ws.cell(
                row=row_no,
                column=3,
            ).number_format = (
                "0.00%"
            )

            ws.cell(
                row=row_no,
                column=5,
            ).number_format = (
                "#,##0.0000"
            )

            ws.cell(
                row=row_no,
                column=6,
            ).number_format = (
                "#,##0.0000"
            )

        for col in range(
            1,
            7,
        ):
            ws.cell(
                row=row_no,
                column=col,
            ).border = THIN_BORDER

        row_no += 1

    # ========================================================
    # MEVCUT 15 GRAFİK
    # YERLERİ VE BOYUTLARI DEĞİŞMEDİ
    # ========================================================

    target_banks = TARGET_BANKS

    product_configs = [
        (
            "USD",
            "DOLAR",
        ),
        (
            "EUR",
            "EURO",
        ),
        (
            "XAU",
            "GRAM ALTIN",
        ),
    ]

    helper_base_row = 200
    helper_gap_rows = 3

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

        run_at = row.get(
            "run_at"
        )

        item_dt = _parse_dt(
            run_at
        )

        if provider not in target_banks:
            continue

        if code not in {
            "USD",
            "EUR",
            "XAU",
        }:
            continue

        if (
            not run_at
            or not item_dt
        ):
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
                    (sell - buy)
                    / buy
                ) * 100.0

        if spread_pct is None:
            continue

        run_bucket = (
            bank_trends.setdefault(
                run_at,
                {
                    "dt": item_dt,
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
        key=lambda item:
        item["dt"],
    )

    helper_block_height = (
        len(trend_runs)
        + 1
        + helper_gap_rows
    )

    # SENİN MEVCUT YERLEŞİMİN
    chart_columns = [
        (
            "A",
            "J",
        ),
        (
            "J",
            "T",
        ),
        (
            "T",
            "AD",
        ),
    ]

    # SENİN MEVCUT DİKEY YERLEŞİMİN
    chart_row_starts = [
        3,
        26,
        49,
        72,
        95,
    ]

    for bank_index, bank in enumerate(
        target_banks
    ):

        helper_start_row = (
            helper_base_row
            + bank_index
            * helper_block_height
        )

        helper_headers = [
            "Çekim Zamanı",
            "DOLAR",
            "EURO",
            "GRAM ALTIN",
        ]

        for col, header in enumerate(
            helper_headers,
            start=1,
        ):

            cell = ws.cell(
                row=helper_start_row,
                column=col,
                value=header,
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="EAF2F8",
            )

            cell.font = Font(
                bold=True,
                color="1F4E78",
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        category_values = []

        cached_values = {
            "USD": [],
            "EUR": [],
            "XAU": [],
        }

        for helper_row, run in enumerate(
            trend_runs,
            start=helper_start_row + 1,
        ):

            item_dt = run["dt"]

            # HAFTA SONU BELİRTECİ
            weekday_suffix = (
                " Cmt"
                if item_dt.weekday()
                == 5
                else (
                    " Paz"
                    if item_dt.weekday()
                    == 6
                    else ""
                )
            )

            # Saat korunuyor, mevcut formatı bozmadım.
            label = (
                item_dt.strftime(
                    "%d.%m"
                )
                + weekday_suffix
                + " "
                + item_dt.strftime(
                    "%H:%M"
                )
            )

            category_values.append(
                label
            )

            ws.cell(
                helper_row,
                1,
                label,
            )

            bank_values = (
                run[
                    "banks"
                ].get(
                    bank,
                    {},
                )
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
                    bank_values.get(
                        code
                    )
                )

                cached_values[
                    code
                ].append(
                    value
                )

                cell = ws.cell(
                    helper_row,
                    col,
                )

                if value is not None:
                    cell.value = value
                    cell.number_format = (
                        "0.00%"
                    )

        if len(trend_runs) < 2:
            continue

        helper_end_row = (
            helper_start_row
            + len(trend_runs)
        )

        for product_index, (
            code,
            product_name,
        ) in enumerate(
            product_configs
        ):

            chart = LineChart()

            chart.style = 10

            chart.title = (
                f"{bank} - "
                f"{product_name} "
                "Makas %"
            )

            chart.y_axis.title = (
                "Makas %"
            )

            chart.x_axis.title = None

            # BOYUTLAR SENİN ÇALIŞAN KODUNLA AYNI
            chart.height = 10.0
            chart.width = 21.0

            try:
                chart.x_axis.axPos = "b"
                chart.x_axis.delete = False
                chart.x_axis.tickLblPos = "low"
                chart.x_axis.tickLblSkip = 1
                chart.x_axis.tickMarkSkip = 1
                chart.x_axis.majorTickMark = "none"
                chart.x_axis.minorTickMark = "none"
            except Exception:
                pass

            try:
                chart.y_axis.majorGridlines = None
                chart.y_axis.majorTickMark = "none"
                chart.y_axis.minorTickMark = "none"
            except Exception:
                pass

            bank_chart_color = (
                BANK_CHART_COLORS.get(
                    bank,
                    DEFAULT_BANK_CHART_COLOR,
                )
            )

            try:
                chart.graphical_properties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=(
                            bank_chart_color
                        ),
                    )
                )
            except Exception:
                pass

            try:
                chart.plot_area.graphicalProperties = (
                    GraphicalProperties(
                        noFill=False,
                        solidFill=(
                            bank_chart_color
                        ),
                    )
                )
            except Exception:
                pass

            chart.legend = None

            data_col = (
                2 + product_index
            )

            data = Reference(
                ws,
                min_col=data_col,
                max_col=data_col,
                min_row=helper_start_row,
                max_row=helper_end_row,
            )

            cats = Reference(
                ws,
                min_col=1,
                min_row=helper_start_row + 1,
                max_row=helper_end_row,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            chart.set_categories(
                cats
            )

            marker_symbol = (
                "circle"
                if code == "USD"
                else (
                    "square"
                    if code == "EUR"
                    else "triangle"
                )
            )

            try:

                series = (
                    chart.series[
                        0
                    ]
                )

                series.marker.symbol = (
                    marker_symbol
                )

                series.marker.size = 6

                series.graphicalProperties.line.width = (
                    20000
                )

                # YENİ:
                # DOLAR MAVİ
                # EURO TURUNCU
                # ALTIN SARI
                series.graphicalProperties.line.solidFill = (
                    PRODUCT_LINE_COLORS[
                        code
                    ]
                )

            except Exception:
                pass

            chart.dLbls = (
                DataLabelList()
            )

            chart.dLbls.showVal = True
            chart.dLbls.numFmt = "0.00%"
            chart.dLbls.dLblPos = "t"
            chart.dLbls.showLegendKey = False
            chart.dLbls.showCatName = False
            chart.dLbls.showSerName = False

            try:
                chart.dLbls.showLeaderLines = (
                    False
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
                            defRPr=(
                                CharacterProperties(
                                    sz=650
                                )
                            )
                        ),
                        endParaRPr=(
                            CharacterProperties(
                                sz=650
                            )
                        ),
                    )
                ],
            )

            y_axis_text = RichText(
                bodyPr=RichTextProperties(),
                p=[
                    Paragraph(
                        pPr=ParagraphProperties(
                            defRPr=(
                                CharacterProperties(
                                    sz=800
                                )
                            )
                        ),
                        endParaRPr=(
                            CharacterProperties(
                                sz=800
                            )
                        ),
                    )
                ],
            )

            data_label_text = RichText(
                bodyPr=RichTextProperties(),
                p=[
                    Paragraph(
                        pPr=ParagraphProperties(
                            defRPr=(
                                CharacterProperties(
                                    sz=700
                                )
                            )
                        ),
                        endParaRPr=(
                            CharacterProperties(
                                sz=700
                            )
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

                chart.dLbls.txPr = (
                    data_label_text
                )
            except Exception:
                pass

            try:

                title_paragraph = (
                    chart.title.tx.rich.p[
                        0
                    ]
                )

                title_paragraph.pPr = (
                    ParagraphProperties(
                        defRPr=(
                            CharacterProperties(
                                sz=1000
                            )
                        )
                    )
                )

                title_paragraph.endParaRPr = (
                    CharacterProperties(
                        sz=1000
                    )
                )

                for run in (
                    title_paragraph.r
                ):
                    run.rPr = (
                        CharacterProperties(
                            sz=1000,
                            b=True,
                        )
                    )

            except Exception:
                pass

            _cache_line_chart(
                chart,
                category_values,
                [
                    (
                        product_name,
                        cached_values[
                            code
                        ],
                    )
                ],
            )

            try:
                chart.y_axis.numFmt = (
                    "0.00%"
                )
            except Exception:
                pass

            chart_values = [
                value
                for value
                in cached_values[
                    code
                ]
                if value is not None
            ]

            if chart_values:

                minimum = min(
                    chart_values
                )

                maximum = max(
                    chart_values
                )

                padding = max(
                    (
                        maximum
                        - minimum
                    )
                    * 0.20,
                    maximum
                    * 0.02,
                    0.00005,
                )

                chart.y_axis.scaling.min = max(
                    0,
                    minimum
                    - padding,
                )

                chart.y_axis.scaling.max = (
                    maximum
                    + padding
                )

            start_col, end_col = (
                chart_columns[
                    product_index
                ]
            )

            start_row = (
                chart_row_starts[
                    bank_index
                ]
            )

            ws.add_chart(
                chart,
                f"{start_col}"
                f"{start_row}",
            )

    # ========================================================
    # AYLIK ORTALAMA TABLOSU
    # MEVCUT YAPIYI BOZMAMAK İÇİN 140. SATIRDAN BAŞLIYOR
    # ========================================================

    monthly_data = (
        _build_monthly_averages(
            history
        )
    )

    months = sorted(
        monthly_data.keys()
    )

    monthly_title_row = 140

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
        horizontal="center",
        vertical="center",
    )

    monthly_header_row = (
        monthly_title_row + 1
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
            monthly_header_row,
            col,
            header,
        )

    _style_header(
        ws,
        monthly_header_row,
        1,
        5,
    )

    ws.cell(
        monthly_header_row,
        3,
    ).fill = USD_HEADER_FILL

    ws.cell(
        monthly_header_row,
        4,
    ).fill = EUR_HEADER_FILL

    ws.cell(
        monthly_header_row,
        5,
    ).fill = XAU_HEADER_FILL

    monthly_row = (
        monthly_header_row
        + 1
    )

    for year, month in months:

        month_label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        for bank in TARGET_BANKS:

            ws.cell(
                monthly_row,
                1,
                month_label,
            )

            bank_cell = ws.cell(
                monthly_row,
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
                (
                    3,
                    "USD",
                ),
                (
                    4,
                    "EUR",
                ),
                (
                    5,
                    "XAU",
                ),
            ):

                value = (
                    monthly_data
                    .get(
                        (
                            year,
                            month,
                        ),
                        {},
                    )
                    .get(
                        bank,
                        {},
                    )
                    .get(
                        code
                    )
                )

                cell = ws.cell(
                    monthly_row,
                    col,
                )

                cell.fill = (
                    _product_fill(
                        code
                    )
                )

                if value is not None:

                    cell.value = (
                        value
                    )

                    cell.number_format = (
                        "0.00%"
                    )

            for col in range(
                1,
                6,
            ):
                ws.cell(
                    monthly_row,
                    col,
                ).border = THIN_BORDER

            monthly_row += 1

    # ========================================================
    # AYLIK GRAFİK YARDIMCI VERİLERİ
    #
    # Mevcut grafik yardımcı tablolarına karışmaması için
    # 500. satırdan başlıyor.
    # ========================================================

    monthly_helper_row = 500

    ws.cell(
        monthly_helper_row,
        1,
        "Ay",
    )

    monthly_column_map = {}

    helper_col = 2

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        monthly_column_map[
            code
        ] = {}

        for bank in TARGET_BANKS:

            monthly_column_map[
                code
            ][bank] = (
                helper_col
            )

            ws.cell(
                monthly_helper_row,
                helper_col,
                bank,
            )

            helper_col += 1

    monthly_labels = []

    for row_offset, (
        year,
        month,
    ) in enumerate(
        months,
        start=1,
    ):

        target_row = (
            monthly_helper_row
            + row_offset
        )

        month_label = (
            f"{MONTH_NAMES[month]} "
            f"{year}"
        )

        monthly_labels.append(
            month_label
        )

        ws.cell(
            target_row,
            1,
            month_label,
        )

        for code in (
            "USD",
            "EUR",
            "XAU",
        ):

            for bank in (
                TARGET_BANKS
            ):

                value = (
                    monthly_data
                    .get(
                        (
                            year,
                            month,
                        ),
                        {},
                    )
                    .get(
                        bank,
                        {},
                    )
                    .get(
                        code
                    )
                )

                col = (
                    monthly_column_map[
                        code
                    ][bank]
                )

                if value is not None:

                    ws.cell(
                        target_row,
                        col,
                        value,
                    ).number_format = (
                        "0.00%"
                    )

    monthly_end_row = (
        monthly_helper_row
        + len(months)
    )

    # ========================================================
    # 3 AYLIK ORTALAMA GRAFİĞİ
    # MEVCUT 15 GRAFİĞİN ALTINDA
    # ========================================================

    if months:

        monthly_chart_positions = [
            "A165",
            "J165",
            "T165",
        ]

        for product_index, (
            code,
            product_name,
        ) in enumerate(
            product_configs
        ):

            chart = LineChart()

            chart.style = 10

            chart.title = (
                f"{product_name} - "
                "Aylık Ortalama Makas %"
            )

            chart.y_axis.title = (
                "Ortalama Makas %"
            )

            chart.x_axis.title = None

            # Mevcut grafiklerle aynı ölçü
            chart.height = 8.5
            chart.width = 21.0

            try:
                chart.x_axis.axPos = "b"
                chart.x_axis.delete = False
                chart.x_axis.tickLblPos = "low"
                chart.x_axis.majorTickMark = (
                    "none"
                )
                chart.x_axis.minorTickMark = (
                    "none"
                )

                chart.y_axis.majorGridlines = (
                    None
                )

                chart.y_axis.majorTickMark = (
                    "none"
                )

                chart.y_axis.numFmt = (
                    "0.00%"
                )
            except Exception:
                pass

            cached_series = []

            for bank in TARGET_BANKS:

                data_col = (
                    monthly_column_map[
                        code
                    ][bank]
                )

                data = Reference(
                    ws,
                    min_col=data_col,
                    max_col=data_col,
                    min_row=monthly_helper_row,
                    max_row=monthly_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = [
                    (
                        monthly_data
                        .get(
                            month_key,
                            {},
                        )
                        .get(
                            bank,
                            {},
                        )
                        .get(
                            code
                        )
                    )
                    for month_key
                    in months
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
                min_row=(
                    monthly_helper_row
                    + 1
                ),
                max_row=monthly_end_row,
            )

            chart.set_categories(
                cats
            )

            # Her banka farklı renk
            for index, series in enumerate(
                chart.series
            ):

                bank = TARGET_BANKS[
                    index
                ]

                try:
                    series.graphicalProperties.line.solidFill = (
                        COMPARISON_BANK_COLORS[
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

            try:
                chart.legend.position = (
                    "b"
                )
                chart.legend.overlay = (
                    False
                )
            except Exception:
                pass

            _cache_line_chart(
                chart,
                monthly_labels,
                cached_series,
            )

            # Grafik ölçeğini mevcut verilere göre ayarla
            all_values = []

            for _, values in cached_series:
                all_values.extend(
                    [
                        value
                        for value
                        in values
                        if value
                        is not None
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
                    )
                    * 0.20,
                    maximum
                    * 0.02,
                    0.00005,
                )

                chart.y_axis.scaling.min = max(
                    0,
                    minimum
                    - padding,
                )

                chart.y_axis.scaling.max = (
                    maximum
                    + padding
                )

            ws.add_chart(
                chart,
                monthly_chart_positions[
                    product_index
                ],
            )

    # ========================================================
    # SÜTUN GENİŞLİKLERİ
    # SENİN MEVCUT DEĞERLERİN KORUNDU
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

    history = read_history(
        history_path
    )

    if not history:
        raise RuntimeError(
            "Excel oluşturmak için "
            "geçmiş veri bulunamadı."
        )

    # --------------------------------------------------------
    # ÖNEMLİ:
    # CSV'nin kendisine dokunmuyoruz.
    #
    # Aynı gün 5 kez çalıştırılmış olsa bile Excel tarafında
    # o günün SADECE İLK çalıştırması kullanılıyor.
    # --------------------------------------------------------

    daily_history = (
        _daily_history(
            history
        )
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

    default_sheet = (
        wb.active
    )

    wb.remove(
        default_sheet
    )

    # 1
    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )

    # 2
    _build_history_sheet(
        wb,
        daily_history,
    )

    # 3
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
