from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# ============================================================
# AYARLAR
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

DARK_BLUE = "1F4E78"
TITLE_BLUE = "D9EAF7"

USD_COLOR = "4472C4"
EUR_COLOR = "ED7D31"
XAU_COLOR = "FFC000"

USD_LIGHT = "DDEBF7"
EUR_LIGHT = "FCE4D6"
XAU_LIGHT = "FFF2CC"

WEEKEND_COLOR = "C65911"
WEEKEND_LIGHT = "FFE699"

BANK_LINE_COLORS = {
    "Garanti BBVA": "70AD47",
    "Akbank": "C00000",
    "Yapıkredi": "7030A0",
    "Ziraat Bankası": "FFC000",
    "İş Bankası": "4472C4",
}

BANK_CHART_COLORS = {
    "Garanti BBVA": "E2F0D9",
    "Akbank": "FCE4E4",
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


# ============================================================
# STİLLER
# ============================================================

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor=DARK_BLUE,
)

HEADER_FONT = Font(
    color="FFFFFF",
    bold=True,
)

TITLE_FILL = PatternFill(
    fill_type="solid",
    fgColor=TITLE_BLUE,
)

USD_FILL = PatternFill(
    fill_type="solid",
    fgColor=USD_LIGHT,
)

EUR_FILL = PatternFill(
    fill_type="solid",
    fgColor=EUR_LIGHT,
)

XAU_FILL = PatternFill(
    fill_type="solid",
    fgColor=XAU_LIGHT,
)

USD_HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor=USD_COLOR,
)

EUR_HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor=EUR_COLOR,
)

XAU_HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor="BF9000",
)

WEEKEND_FILL = PatternFill(
    fill_type="solid",
    fgColor=WEEKEND_LIGHT,
)

OK_FILL = PatternFill(
    fill_type="solid",
    fgColor="E2F0D9",
)

CONTROL_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFF2CC",
)

ERROR_FILL = PatternFill(
    fill_type="solid",
    fgColor="FCE4D6",
)

THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)

STATUS_LABELS = {
    "OK": "DOĞRU",
    "KONTROL": "KONTROL GEREKLİ",
    "ERROR": "HATA",
}


# ============================================================
# VERİ YARDIMCILARI
# ============================================================

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
        "%Y-%m-%dT%H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    return None


def _to_float(value):

    if value is None or value == "":
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = (
        str(value)
        .strip()
        .replace("%", "")
        .replace("\u00a0", "")
    )

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


def _normalize_code(value):

    text = str(value or "").strip().upper()

    replacements = {
        "DOLAR": "USD",
        "USD": "USD",
        "AMERIKAN DOLARI": "USD",
        "AMERİKAN DOLARI": "USD",

        "EURO": "EUR",
        "EUR": "EUR",

        "XAU": "XAU",
        "ALTIN": "XAU",
        "GRAM ALTIN": "XAU",
        "GRAM_ALTIN": "XAU",
    }

    return replacements.get(text, text)


def _normalize_provider(value):

    text = str(value or "").strip()

    folded = (
        text.casefold()
        .replace("ı", "i")
        .replace("ş", "s")
        .replace("ğ", "g")
        .replace("ü", "u")
        .replace("ö", "o")
        .replace("ç", "c")
    )

    if "garanti" in folded:
        return "Garanti BBVA"

    if "akbank" in folded:
        return "Akbank"

    if "yapi kredi" in folded or "yapikredi" in folded:
        return "Yapıkredi"

    if "ziraat" in folded and "katilim" not in folded and "dinamik" not in folded:
        return "Ziraat Bankası"

    if "is bankasi" in folded or "turkiye is bankasi" in folded:
        return "İş Bankası"

    return text


def _get_value(row, *keys):

    for key in keys:

        if key in row and row.get(key) not in (None, ""):
            return row.get(key)

    return None


def _row_provider(row):

    return _normalize_provider(
        _get_value(
            row,
            "provider",
            "bank",
            "banka",
            "name",
            "provider_name",
        )
    )


def _row_code(row):

    return _normalize_code(
        _get_value(
            row,
            "code",
            "product",
            "currency",
            "symbol",
            "urun",
        )
    )


def _row_run_at(row):

    return _get_value(
        row,
        "run_at",
        "timestamp",
        "datetime",
        "date",
    )


def _get_spread_values(row):

    buy = _to_float(
        _get_value(
            row,
            "buy",
            "alis",
            "buying",
        )
    )

    sell = _to_float(
        _get_value(
            row,
            "sell",
            "satis",
            "selling",
        )
    )

    spread = _to_float(
        _get_value(
            row,
            "spread",
            "makas",
        )
    )

    pct = _to_float(
        _get_value(
            row,
            "spread_pct",
            "spread_percent",
            "makas_pct",
            "makas_yuzde",
        )
    )

    if spread is None and buy is not None and sell is not None:
        spread = sell - buy

    if (
        pct is None
        and spread is not None
        and buy not in (None, 0)
    ):
        pct = (spread / buy) * 100

    return buy, sell, spread, pct


def _provider_fill(provider):

    provider = _normalize_provider(provider)

    return PatternFill(
        fill_type="solid",
        fgColor=PROVIDER_COLORS.get(
            provider,
            DEFAULT_PROVIDER_COLOR,
        ),
    )


def _product_fill(code):

    code = _normalize_code(code)

    if code == "USD":
        return USD_FILL

    if code == "EUR":
        return EUR_FILL

    if code == "XAU":
        return XAU_FILL

    return PatternFill(fill_type=None)


def read_history(path):

    path = Path(path)

    if not path.exists():
        return []

    with path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as f:
        return list(csv.DictReader(f))


# ============================================================
# GÜNDE SADECE İLK RUN
# ============================================================

def _daily_history(history):

    runs_by_day = {}

    for row in history:

        raw = _row_run_at(row)
        dt = _parse_dt(raw)

        if not dt:
            continue

        day = dt.date()

        if day not in runs_by_day:
            runs_by_day[day] = dt
        elif dt < runs_by_day[day]:
            runs_by_day[day] = dt

    selected = set(runs_by_day.values())

    result = []

    for row in history:

        dt = _parse_dt(_row_run_at(row))

        if dt in selected:
            result.append(row)

    return result


def _latest_run_rows(history):

    valid = []

    for row in history:

        dt = _parse_dt(_row_run_at(row))

        if dt:
            valid.append((dt, row))

    if not valid:
        return None, []

    latest_dt = max(dt for dt, _ in valid)

    rows = [
        row
        for dt, row in valid
        if dt == latest_dt
    ]

    return latest_dt, rows


# ============================================================
# EXCEL STİL YARDIMCILARI
# ============================================================

def _style_header(
    ws,
    row,
    start_col,
    end_col,
):

    for col in range(
        start_col,
        end_col + 1,
    ):

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


def _apply_table(
    ws,
    start_row,
    end_row,
    end_col,
    name,
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


# ============================================================
# TARİH ETİKETİ
# ============================================================

def _date_label(dt):

    if dt.weekday() == 5:
        return dt.strftime("%d.%m") + " Cmt"

    if dt.weekday() == 6:
        return dt.strftime("%d.%m") + " Paz"

    return dt.strftime("%d.%m")


# ============================================================
# GÜNCEL KURLAR
# ============================================================

def _build_current_sheet(
    wb,
    latest_dt,
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
        15,
    )

    for col in range(4, 8):
        ws.cell(1, col).fill = USD_HEADER_FILL

    for col in range(8, 12):
        ws.cell(1, col).fill = EUR_HEADER_FILL

    for col in range(12, 16):
        ws.cell(1, col).fill = XAU_HEADER_FILL

    providers = {}

    for row in latest_rows:

        provider = _row_provider(row)
        code = _row_code(row)

        if not provider or code not in PRODUCT_ORDER:
            continue

        providers.setdefault(
            provider,
            {},
        )[code] = row

    for row_no, provider in enumerate(
        sorted(
            providers.keys(),
            key=str.casefold,
        ),
        start=2,
    ):

        ws.cell(
            row_no,
            1,
            latest_dt.date(),
        )

        ws.cell(
            row_no,
            2,
            latest_dt.time(),
        )

        ws.cell(
            row_no,
            3,
            provider,
        )

        ws.cell(
            row_no,
            3,
        ).fill = _provider_fill(provider)

        ws.cell(
            row_no,
            3,
        ).font = Font(bold=True)

        layout = {
            "USD": (4, 5, 6, 7),
            "EUR": (8, 9, 10, 11),
            "XAU": (12, 13, 14, 15),
        }

        for code, columns in layout.items():

            for col in columns:
                ws.cell(
                    row_no,
                    col,
                ).fill = _product_fill(code)

            item = providers[provider].get(code)

            if not item:
                continue

            buy, sell, spread, pct = _get_spread_values(item)

            values = (
                buy,
                sell,
                spread,
                pct / 100 if pct is not None else None,
            )

            for col, value in zip(
                columns,
                values,
            ):

                if value is not None:
                    ws.cell(
                        row_no,
                        col,
                        value,
                    )

        ws.cell(
            row_no,
            1,
        ).number_format = "dd.mm.yyyy"

        ws.cell(
            row_no,
            2,
        ).number_format = "hh:mm:ss"

        for col in (
            4, 5, 6,
            8, 9, 10,
            12, 13, 14,
        ):
            ws.cell(
                row_no,
                col,
            ).number_format = "#,##0.0000"

        for col in (7, 11, 15):
            ws.cell(
                row_no,
                col,
            ).number_format = "0.00%"

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
            "C": 25,

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

def _history_sort_key(row):

    dt = (
        _parse_dt(_row_run_at(row))
        or datetime.min
    )

    return (
        dt,
        _row_provider(row).casefold(),
        PRODUCT_ORDER.get(
            _row_code(row),
            99,
        ),
    )


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

    for row_no, item in enumerate(
        sorted(
            history,
            key=_history_sort_key,
        ),
        start=2,
    ):

        dt = _parse_dt(
            _row_run_at(item)
        )

        provider = _row_provider(item)
        code = _row_code(item)

        buy, sell, spread, pct = (
            _get_spread_values(item)
        )

        scraped = _parse_dt(
            _get_value(
                item,
                "scraped_at",
                "scrape_time",
            )
        )

        site_spread = _to_float(
            item.get("site_spread")
        )

        site_pct = _to_float(
            item.get("site_spread_pct")
        )

        if dt:

            ws.cell(
                row_no,
                1,
                dt.date(),
            )

            ws.cell(
                row_no,
                2,
                dt.time(),
            )

        ws.cell(
            row_no,
            3,
            provider,
        )

        ws.cell(
            row_no,
            3,
        ).fill = _provider_fill(provider)

        ws.cell(
            row_no,
            3,
        ).font = Font(bold=True)

        ws.cell(
            row_no,
            4,
            PRODUCT_NAMES.get(
                code,
                code,
            ),
        )

        for col in range(4, 9):
            ws.cell(
                row_no,
                col,
            ).fill = _product_fill(code)

        if buy is not None:
            ws.cell(row_no, 5, buy)

        if sell is not None:
            ws.cell(row_no, 6, sell)

        if spread is not None:
            ws.cell(row_no, 7, spread)

        if pct is not None:
            ws.cell(
                row_no,
                8,
                pct / 100,
            )

        status = (
            item.get("status")
            or "OK"
        )

        ws.cell(
            row_no,
            9,
            STATUS_LABELS.get(
                status,
                status,
            ),
        )

        source = (
            item.get("source_url")
            or item.get("source")
            or ""
        )

        ws.cell(
            row_no,
            10,
            source,
        )

        if source:
            ws.cell(
                row_no,
                10,
            ).hyperlink = source

            ws.cell(
                row_no,
                10,
            ).style = "Hyperlink"

        if site_spread is not None:
            ws.cell(
                row_no,
                11,
                site_spread,
            )

        if site_pct is not None:
            ws.cell(
                row_no,
                12,
                site_pct / 100,
            )

        ws.cell(
            row_no,
            13,
            item.get("note", ""),
        )

        if scraped:
            ws.cell(
                row_no,
                14,
                scraped.time(),
            )

        if dt and dt.weekday() in (5, 6):

            ws.cell(
                row_no,
                1,
            ).fill = WEEKEND_FILL

            ws.cell(
                row_no,
                1,
            ).font = Font(
                bold=True,
                color=WEEKEND_COLOR,
            )

        if status == "ERROR":

            ws.cell(
                row_no,
                9,
            ).fill = ERROR_FILL

        elif status == "KONTROL":

            ws.cell(
                row_no,
                9,
            ).fill = CONTROL_FILL

        else:

            ws.cell(
                row_no,
                9,
            ).fill = OK_FILL

        ws.cell(
            row_no,
            9,
        ).font = Font(bold=True)

        ws.cell(
            row_no,
            1,
        ).number_format = "dd.mm.yyyy"

        ws.cell(
            row_no,
            2,
        ).number_format = "hh:mm:ss"

        ws.cell(
            row_no,
            14,
        ).number_format = "hh:mm:ss"

        for col in (5, 6, 7, 11):
            ws.cell(
                row_no,
                col,
            ).number_format = "#,##0.0000"

        for col in (8, 12):
            ws.cell(
                row_no,
                col,
            ).number_format = "0.00%"

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
            "C": 25,
            "D": 16,
            "E": 15,
            "F": 15,
            "G": 15,
            "H": 14,
            "I": 20,
            "J": 45,
            "K": 17,
            "L": 18,
            "M": 55,
            "N": 22,
        },
    )


# ============================================================
# TREND VERİLERİ
# ============================================================

def _build_bank_trends(history):

    data = {}

    for row in history:

        bank = _row_provider(row)
        code = _row_code(row)

        if bank not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if row.get("status") == "ERROR":
            continue

        dt = _parse_dt(
            _row_run_at(row)
        )

        if not dt:
            continue

        _, _, _, pct = (
            _get_spread_values(row)
        )

        if pct is None:
            continue

        key = dt

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

    return [
        data[key]
        for key in sorted(data)
    ]


# ============================================================
# AYLIK ORTALAMA
# ============================================================

def _monthly_averages(history):

    buckets = {}

    for row in history:

        bank = _row_provider(row)
        code = _row_code(row)

        if bank not in TARGET_BANKS:
            continue

        if code not in PRODUCT_ORDER:
            continue

        if row.get("status") == "ERROR":
            continue

        dt = _parse_dt(
            _row_run_at(row)
        )

        if not dt:
            continue

        _, _, _, pct = (
            _get_spread_values(row)
        )

        if pct is None:
            continue

        key = (
            dt.year,
            dt.month,
            bank,
            code,
        )

        buckets.setdefault(
            key,
            [],
        ).append(
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
        )[code] = (
            sum(values) / len(values)
        )

    return result


# ============================================================
# GRAFİK HELPER SAYFASI
# ============================================================

def _build_chart_data_sheet(
    wb,
    trends,
    monthly,
):

    ws = wb.create_sheet(
        "GRAFIK_VERILERI"
    )

    # --------------------------------------------------------
    # GÜNLÜK
    # --------------------------------------------------------

    ws["A1"] = "Tarih"

    daily_columns = {}

    col = 2

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

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
                        row_no,
                        daily_columns[code][bank],
                        value,
                    )

                    cell.number_format = "0.00%"

    # --------------------------------------------------------
    # AYLIK
    # --------------------------------------------------------

    months = sorted(
        monthly.keys()
    )

    monthly_date_col = 20

    ws.cell(
        1,
        monthly_date_col,
        "Ay",
    )

    monthly_columns = {}

    col = monthly_date_col + 1

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        monthly_columns[code] = {}

        for bank in TARGET_BANKS:

            monthly_columns[code][bank] = col

            ws.cell(
                1,
                col,
                f"{code} - {bank}",
            )

            col += 1

    for row_no, (
        year,
        month,
    ) in enumerate(
        months,
        start=2,
    ):

        ws.cell(
            row_no,
            monthly_date_col,
            f"{MONTH_NAMES[month]} {year}",
        )

        for code in (
            "USD",
            "EUR",
            "XAU",
        ):

            for bank in TARGET_BANKS:

                value = (
                    monthly
                    .get((year, month), {})
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    cell = ws.cell(
                        row_no,
                        monthly_columns[code][bank],
                        value,
                    )

                    cell.number_format = "0.00%"

    return (
        ws,
        daily_columns,
        monthly_columns,
        monthly_date_col,
        months,
    )


# ============================================================
# GRAFİK STİLİ
# ============================================================

def _style_chart(
    chart,
    title,
    legend=True,
    background=None,
):

    chart.title = title

    # Üç grafik yan yana sığacak şekilde.
    chart.width = 16.8
    chart.height = 8.2

    chart.y_axis.title = "Makas %"
    chart.y_axis.numFmt = "0.00%"

    # --------------------------------------------------------
    # GRID ÇİZGİLERİ KAPALI
    # --------------------------------------------------------

    chart.y_axis.majorGridlines = None
    chart.x_axis.majorGridlines = None

    chart.y_axis.majorTickMark = "none"
    chart.x_axis.majorTickMark = "none"

    # Dış çerçeve çok hafif.
    try:

        chart.graphical_properties = GraphicalProperties(
            noFill=True
        )

    except Exception:
        pass

    # Bankaya özel açık arka plan.
    if background:

        try:

            chart.plot_area.graphicalProperties = (
                GraphicalProperties(
                    solidFill=background
                )
            )

        except Exception:
            pass

    # Tarih etiketleri
    try:

        chart.x_axis.tickLblPos = "low"

        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(
                rot=-2700000
            ),
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

    except Exception:
        pass

    if legend:

        chart.legend.position = "b"
        chart.legend.overlay = False

    else:

        chart.legend = None


def _set_series_color(
    series,
    color,
):

    try:

        series.graphicalProperties.line.solidFill = color

        series.graphicalProperties.line.width = 22000

        series.marker.symbol = "circle"
        series.marker.size = 6

        series.marker.graphicalProperties.solidFill = color

        series.marker.graphicalProperties.line.solidFill = color

        series.smooth = False

    except Exception:
        pass


def _set_axis_bounds(
    chart,
    values,
):

    valid = [
        value
        for value in values
        if value is not None
    ]

    if not valid:
        return

    minimum = min(valid)
    maximum = max(valid)

    difference = maximum - minimum

    padding = max(
        difference * 0.20,
        maximum * 0.03,
        0.0005,
    )

    chart.y_axis.scaling.min = max(
        0,
        minimum - padding,
    )

    chart.y_axis.scaling.max = (
        maximum + padding
    )


# ============================================================
# ÜST TABLOLAR
# ============================================================

def _build_latest_table(
    ws,
    latest_dt,
    latest_rows,
):

    # --------------------------------------------------------
    # SOL
    # A3:D10
    # --------------------------------------------------------

    ws.merge_cells(
        "A3:D3"
    )

    ws["A3"] = "SON VERİ ÇEKİMİ"

    ws["A3"].fill = TITLE_FILL
    ws["A3"].font = Font(
        bold=True,
        color=DARK_BLUE,
    )
    ws["A3"].alignment = Alignment(
        horizontal="center"
    )

    providers = {
        _row_provider(row)
        for row in latest_rows
        if _row_provider(row)
    }

    info = [
        (
            "Veri tarihi",
            latest_dt.date(),
        ),
        (
            "Çekim saati",
            latest_dt.time(),
        ),
        (
            "Sağlayıcı sayısı",
            len(providers),
        ),
        (
            "Ürün kaydı",
            len(latest_rows),
        ),
        (
            "Hatalı kayıt",
            sum(
                1
                for row in latest_rows
                if row.get("status") == "ERROR"
            ),
        ),
        (
            "Kontrol gereken",
            sum(
                1
                for row in latest_rows
                if row.get("status") == "KONTROL"
            ),
        ),
    ]

    row_no = 4

    for label, value in info:

        ws.merge_cells(
            start_row=row_no,
            start_column=1,
            end_row=row_no,
            end_column=2,
        )

        ws.cell(
            row_no,
            1,
            label,
        )

        ws.merge_cells(
            start_row=row_no,
            start_column=3,
            end_row=row_no,
            end_column=4,
        )

        ws.cell(
            row_no,
            3,
            value,
        )

        ws.cell(
            row_no,
            1,
        ).font = Font(bold=True)

        ws.cell(
            row_no,
            1,
        ).fill = PatternFill(
            fill_type="solid",
            fgColor="EAF2F8",
        )

        ws.cell(
            row_no,
            1,
        ).alignment = Alignment(
            vertical="center"
        )

        ws.cell(
            row_no,
            3,
        ).alignment = Alignment(
            horizontal="center"
        )

        row_no += 1

    ws["C4"].number_format = "dd.mm.yyyy"
    ws["C5"].number_format = "hh:mm:ss"


def _build_monthly_table(
    ws,
    monthly,
    months,
):

    # --------------------------------------------------------
    # ORTA
    # F3:J...
    # --------------------------------------------------------

    ws.merge_cells(
        "F3:J3"
    )

    ws["F3"] = "AYLIK ORTALAMA MAKAS %"

    ws["F3"].fill = TITLE_FILL
    ws["F3"].font = Font(
        bold=True,
        color=DARK_BLUE,
    )
    ws["F3"].alignment = Alignment(
        horizontal="center"
    )

    headers = [
        "Ay",
        "Banka",
        "Dolar",
        "Euro",
        "Gram Altın",
    ]

    for col, value in enumerate(
        headers,
        start=6,
    ):

        ws.cell(
            4,
            col,
            value,
        )

    _style_header(
        ws,
        4,
        6,
        10,
    )

    ws["H4"].fill = USD_HEADER_FILL
    ws["I4"].fill = EUR_HEADER_FILL
    ws["J4"].fill = XAU_HEADER_FILL

    row_no = 5

    for year, month in months:

        month_label = (
            f"{MONTH_NAMES[month]} {year}"
        )

        for bank in TARGET_BANKS:

            ws.cell(
                row_no,
                6,
                month_label,
            )

            ws.cell(
                row_no,
                7,
                bank,
            )

            ws.cell(
                row_no,
                7,
            ).fill = _provider_fill(bank)

            ws.cell(
                row_no,
                7,
            ).font = Font(bold=True)

            for col, code in (
                (8, "USD"),
                (9, "EUR"),
                (10, "XAU"),
            ):

                cell = ws.cell(
                    row_no,
                    col,
                )

                cell.fill = _product_fill(code)

                value = (
                    monthly
                    .get((year, month), {})
                    .get(bank, {})
                    .get(code)
                )

                if value is not None:

                    cell.value = value
                    cell.number_format = "0.00%"

            for col in range(6, 11):

                ws.cell(
                    row_no,
                    col,
                ).border = THIN_BORDER

            row_no += 1

    return row_no - 1


def _build_best_table(
    ws,
    latest_rows,
):

    # --------------------------------------------------------
    # SAĞ
    # L3:Q8
    # --------------------------------------------------------

    ws.merge_cells(
        "L3:Q3"
    )

    ws["L3"] = (
        "5 BANKADA EN DÜŞÜK GÜNCEL MAKAS"
    )

    ws["L3"].fill = TITLE_FILL
    ws["L3"].font = Font(
        bold=True,
        color=DARK_BLUE,
    )
    ws["L3"].alignment = Alignment(
        horizontal="center"
    )

    headers = [
        "Ürün",
        "Banka Sayısı",
        "En Düşük %",
        "En Avantajlı",
        "Alış",
        "Satış",
    ]

    for col, value in enumerate(
        headers,
        start=12,
    ):

        ws.cell(
            4,
            col,
            value,
        )

    _style_header(
        ws,
        4,
        12,
        17,
    )

    row_no = 5

    for code in (
        "USD",
        "EUR",
        "XAU",
    ):

        candidates = []

        found_banks = set()

        for item in latest_rows:

            bank = _row_provider(item)
            item_code = _row_code(item)

            if bank not in TARGET_BANKS:
                continue

            if item_code != code:
                continue

            if item.get("status") == "ERROR":
                continue

            buy, sell, _, pct = (
                _get_spread_values(item)
            )

            if (
                buy is None
                or sell is None
                or pct is None
            ):
                continue

            found_banks.add(bank)

            candidates.append(
                (
                    pct,
                    bank,
                    buy,
                    sell,
                )
            )

        best = (
            min(
                candidates,
                key=lambda x: x[0],
            )
            if candidates
            else None
        )

        ws.cell(
            row_no,
            12,
            PRODUCT_NAMES[code],
        )

        ws.cell(
            row_no,
            12,
        ).fill = _product_fill(code)

        ws.cell(
            row_no,
            12,
        ).font = Font(bold=True)

        ws.cell(
            row_no,
            13,
            len(found_banks),
        )

        if best:

            pct, bank, buy, sell = best

            ws.cell(
                row_no,
                14,
                pct / 100,
            )

            ws.cell(
                row_no,
                14,
            ).number_format = "0.00%"

            ws.cell(
                row_no,
                15,
                bank,
            )

            ws.cell(
                row_no,
                15,
            ).fill = _provider_fill(bank)

            ws.cell(
                row_no,
                15,
            ).font = Font(bold=True)

            ws.cell(
                row_no,
                16,
                buy,
            )

            ws.cell(
                row_no,
                17,
                sell,
            )

            ws.cell(
                row_no,
                16,
            ).number_format = "#,##0.0000"

            ws.cell(
                row_no,
                17,
            ).number_format = "#,##0.0000"

        for col in range(12, 18):

            ws.cell(
                row_no,
                col,
            ).border = THIN_BORDER

        row_no += 1


# ============================================================
# OZET
# ============================================================

def _build_summary_sheet(
    wb,
    latest_dt,
    latest_rows,
    history,
):

    ws = wb.create_sheet("OZET")

    ws.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # VERİLER
    # --------------------------------------------------------

    trends = _build_bank_trends(
        history
    )

    monthly = _monthly_averages(
        history
    )

    (
        data_ws,
        daily_columns,
        monthly_columns,
        monthly_date_col,
        months,
    ) = _build_chart_data_sheet(
        wb,
        trends,
        monthly,
    )

    # --------------------------------------------------------
    # ANA BAŞLIK
    # --------------------------------------------------------

    ws.merge_cells(
        "A1:Q1"
    )

    ws["A1"] = (
        "Döviz ve Altın Kur Takip Özeti"
    )

    ws["A1"].fill = TITLE_FILL

    ws["A1"].font = Font(
        bold=True,
        size=16,
        color=DARK_BLUE,
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    ws.row_dimensions[1].height = 26

    # --------------------------------------------------------
    # ÜST TABLOLAR
    #
    # SOL: SON VERİ
    # ORTA: AYLIK ORTALAMA
    # SAĞ: EN DÜŞÜK MAKAS
    # --------------------------------------------------------

    _build_latest_table(
        ws,
        latest_dt,
        latest_rows,
    )

    monthly_last_row = (
        _build_monthly_table(
            ws,
            monthly,
            months,
        )
    )

    _build_best_table(
        ws,
        latest_rows,
    )

    # --------------------------------------------------------
    # TABLOLARIN BİTTİĞİ SATIR
    # --------------------------------------------------------

    top_tables_end = max(
        10,
        monthly_last_row,
    )

    # Tablolarla grafik arasında net boşluk
    DAILY_CHART_ROW = (
        top_tables_end + 4
    )

    MONTHLY_CHART_ROW = (
        DAILY_CHART_ROW + 20
    )

    INDIVIDUAL_START_ROW = (
        MONTHLY_CHART_ROW + 20
    )

    # --------------------------------------------------------
    # GRAFİK KOLONLARI
    #
    # A / G / M
    #
    # Önceki J/S yerleşiminden daha dengeli.
    # --------------------------------------------------------

    CHART_POSITIONS = [
        "A",
        "G",
        "M",
    ]

    daily_end_row = (
        len(trends) + 1
    )

    # ========================================================
    # 1. SATIR GRAFİKLER:
    # 5 BANKA TOPLU
    # ========================================================

    if trends:

        daily_categories = Reference(
            data_ws,
            min_col=1,
            min_row=2,
            max_row=daily_end_row,
        )

        for product_index, code in enumerate(
            ("USD", "EUR", "XAU")
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{PRODUCT_NAMES[code]} "
                    f"- 5 Banka Makas %"
                ),
                legend=True,
            )

            all_values = []

            for bank in TARGET_BANKS:

                col = (
                    daily_columns[code][bank]
                )

                data = Reference(
                    data_ws,
                    min_col=col,
                    max_col=col,
                    min_row=1,
                    max_row=daily_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = [
                    run["banks"]
                    .get(bank, {})
                    .get(code)
                    for run in trends
                ]

                all_values.extend(
                    value
                    for value in values
                    if value is not None
                )

            chart.set_categories(
                daily_categories
            )

            for index, series in enumerate(
                chart.series
            ):

                bank = TARGET_BANKS[index]

                _set_series_color(
                    series,
                    BANK_LINE_COLORS[bank],
                )

            _set_axis_bounds(
                chart,
                all_values,
            )

            ws.add_chart(
                chart,
                (
                    f"{CHART_POSITIONS[product_index]}"
                    f"{DAILY_CHART_ROW}"
                ),
            )

    # ========================================================
    # 2. SATIR GRAFİKLER:
    # AYLIK ORTALAMA
    # ========================================================

    if months:

        monthly_end_row = (
            len(months) + 1
        )

        monthly_categories = Reference(
            data_ws,
            min_col=monthly_date_col,
            min_row=2,
            max_row=monthly_end_row,
        )

        for product_index, code in enumerate(
            ("USD", "EUR", "XAU")
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{PRODUCT_NAMES[code]} "
                    f"- Aylık Ortalama Makas %"
                ),
                legend=True,
            )

            all_values = []

            for bank in TARGET_BANKS:

                col = (
                    monthly_columns[code][bank]
                )

                data = Reference(
                    data_ws,
                    min_col=col,
                    max_col=col,
                    min_row=1,
                    max_row=monthly_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = [
                    monthly
                    .get(month, {})
                    .get(bank, {})
                    .get(code)
                    for month in months
                ]

                all_values.extend(
                    value
                    for value in values
                    if value is not None
                )

            chart.set_categories(
                monthly_categories
            )

            for index, series in enumerate(
                chart.series
            ):

                bank = TARGET_BANKS[index]

                # ÇİZGİ + NOKTA AYNI RENK
                _set_series_color(
                    series,
                    BANK_LINE_COLORS[bank],
                )

            _set_axis_bounds(
                chart,
                all_values,
            )

            ws.add_chart(
                chart,
                (
                    f"{CHART_POSITIONS[product_index]}"
                    f"{MONTHLY_CHART_ROW}"
                ),
            )

    # ========================================================
    # TEKİL BANKA GRAFİKLERİ
    #
    # 5 BANKA × 3 ÜRÜN
    # ========================================================

    for bank_index, bank in enumerate(
        TARGET_BANKS
    ):

        bank_row = (
            INDIVIDUAL_START_ROW
            + bank_index * 20
        )

        for product_index, code in enumerate(
            ("USD", "EUR", "XAU")
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{bank} - "
                    f"{PRODUCT_NAMES[code]} Makas %"
                ),
                legend=False,
                background=(
                    BANK_CHART_COLORS[bank]
                ),
            )

            if not trends:
                continue

            col = (
                daily_columns[code][bank]
            )

            data = Reference(
                data_ws,
                min_col=col,
                max_col=col,
                min_row=1,
                max_row=daily_end_row,
            )

            chart.add_data(
                data,
                titles_from_data=True,
            )

            daily_categories = Reference(
                data_ws,
                min_col=1,
                min_row=2,
                max_row=daily_end_row,
            )

            chart.set_categories(
                daily_categories
            )

            # ------------------------------------------------
            # TEK BANKA GRAFİĞİ:
            #
            # Dolar = mavi
            # Euro = turuncu
            # Altın = sarı
            #
            # NOKTA DA AYNI RENK
            # ------------------------------------------------

            if chart.series:

                color = {
                    "USD": USD_COLOR,
                    "EUR": EUR_COLOR,
                    "XAU": XAU_COLOR,
                }[code]

                _set_series_color(
                    chart.series[0],
                    color,
                )

            values = [
                run["banks"]
                .get(bank, {})
                .get(code)
                for run in trends
            ]

            _set_axis_bounds(
                chart,
                values,
            )

            # ------------------------------------------------
            # VERİ ETİKETİ YOK
            #
            # Böylece:
            #
            # USD - Garanti BBVA; 28.08 ...
            #
            # gibi yazılar grafiğin içine dolmaz.
            # ------------------------------------------------

            chart.dLbls = None

            ws.add_chart(
                chart,
                (
                    f"{CHART_POSITIONS[product_index]}"
                    f"{bank_row}"
                ),
            )

    # --------------------------------------------------------
    # OZET SÜTUN GENİŞLİKLERİ
    # --------------------------------------------------------

    _set_widths(
        ws,
        {
            "A": 17,
            "B": 17,
            "C": 15,
            "D": 15,

            "E": 3,

            "F": 17,
            "G": 20,
            "H": 15,
            "I": 15,
            "J": 17,

            "K": 3,

            "L": 17,
            "M": 14,
            "N": 15,
            "O": 19,
            "P": 16,
            "Q": 16,

            "R": 3,

            "S": 16,
            "T": 16,
            "U": 16,
            "V": 16,
            "W": 16,
            "X": 16,
        },
    )

    # --------------------------------------------------------
    # GRAFİK HELPER SAYFASI GİZLİ
    # --------------------------------------------------------

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

    latest_dt, latest_rows = (
        _latest_run_rows(
            raw_history
        )
    )

    if (
        latest_dt is None
        or not latest_rows
    ):

        raise RuntimeError(
            "Son çekime ait veri bulunamadı."
        )

    # --------------------------------------------------------
    # GEÇMİŞ + GRAFİKLER:
    # AYNI GÜNDE SADECE İLK RUN
    # --------------------------------------------------------

    daily_history = (
        _daily_history(
            raw_history
        )
    )

    if not daily_history:

        raise RuntimeError(
            "Günlük geçmiş veri bulunamadı."
        )

    wb = Workbook()

    default_ws = wb.active

    wb.remove(
        default_ws
    )

    # --------------------------------------------------------
    # 1 - GÜNCEL
    # --------------------------------------------------------

    _build_current_sheet(
        wb,
        latest_dt,
        latest_rows,
    )

    # --------------------------------------------------------
    # 2 - GEÇMİŞ
    # --------------------------------------------------------

    _build_history_sheet(
        wb,
        daily_history,
    )

    # --------------------------------------------------------
    # 3 - ÖZET
    # --------------------------------------------------------

    _build_summary_sheet(
        wb,
        latest_dt,
        latest_rows,
        daily_history,
    )

    # --------------------------------------------------------
    # İLK AÇILAN SAYFA
    # --------------------------------------------------------

    wb.active = wb.sheetnames.index(
        "OZET"
    )

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
# ESKİ main.py UYUMLULUĞU
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
