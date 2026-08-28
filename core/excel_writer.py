from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


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

PROVIDER_COLORS = {
    'Akbank': 'ECCACA',
    'Albaraka Türk': 'DFF1E4',
    'Alternatif Bank': 'DECDEA',
    'Altınkaynak': 'F3F0DD',
    'Anadolubank': 'CDE5EA',
    'CEPTETEB': 'F1DFE9',
    'Denizbank': 'D3ECCA',
    'DestekBank': 'E0DFF1',
    'Dünya Katılım': 'EAD7CD',
    'Emlak Katılım': 'DDF3EB',
    'Enpara': 'E7CDEA',
    'Fibabanka': 'EDF1DF',
    'Garanti BBVA': 'CADBEC',
    'Getirfinans': 'F1DFE3',
    'Halkbank': 'CDEACF',
    'Harem': 'E5DDF3',
    'Hayat Finans': 'EAE0CD',
    'Hepsipay': 'DFF1F1',
    'HSBC': 'ECCAE4',
    'ING Bank': 'E7F1DF',
    'İş Bankası': 'CDD2EA',
    'Kapalıçarşı': 'F3E0DD',
    'Kuveyt Türk': 'CDEAD9',
    'Merkez Bankası': 'ECDFF1',
    'Misyon Bank': 'ECECCA',
    'Odacı': 'DFECF1',
    'Odeabank': 'EACDD9',
    'Papara': 'E0F3DD',
    'QNB Finansbank': 'D2CDEA',
    'TOM Bank Hadi': 'F1E7DF',
    'Türkiye Finans': 'CAECE4',
    'Vakıf Katılım': 'F1DFF0',
    'Vakıfbank': 'E0EACD',
    'Venüs': 'DDE5F3',
    'Yapıkredi': 'EACDCF',
    'Ziraat Bankası': 'DFF1E3',
    'Ziraat Dinamik': 'DBCAEC',
    'Ziraat Katılım': 'F1EDDF',
}

DEFAULT_PROVIDER_COLOR = 'E8EDF3'

PRODUCT_ORDER = {
    "USD": 0,
    "EUR": 1,
    "XAU": 2,
}



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

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _style_header(ws, row: int, start_col: int, end_col: int) -> None:
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
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


def _display_status(status: str | None) -> str:
    raw = (status or "").strip()
    return STATUS_LABELS.get(raw, raw)


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
        key=lambda row: _parse_dt(row.get("run_at")),
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


def _provider_map(
    rows: Iterable[dict],
) -> dict[str, dict[str, dict]]:
    result: dict[str, dict[str, dict]] = {}

    for row in rows:
        provider = row.get("provider", "").strip()
        code = row.get("code", "").strip()

        if not provider or not code:
            continue

        result.setdefault(provider, {})[code] = row

    return result


def _history_sort_key(row: dict):
    provider = (row.get("provider") or "").strip().casefold()
    code = (row.get("code") or "").strip()
    run_dt = _parse_dt(row.get("run_at")) or datetime.min

    return (
        provider,
        PRODUCT_ORDER.get(code, 99),
        run_dt,
    )


def _build_trend_data(
    history: list[dict],
) -> list[tuple[datetime, dict[str, float | None]]]:
    """
    Her çekim için DOLAR / EURO / GRAM ALTIN'daki
    en düşük geçerli makas yüzdesini bulur.

    CSV'deki spread_pct yüzde birimindedir.
    Örnek: 0.02 = %0,02. Excel'e yazarken /100 yapılır.
    """
    runs: dict[str, dict] = {}

    for row in history:
        run_at = row.get("run_at")
        run_dt = _parse_dt(run_at)

        if not run_at or not run_dt:
            continue

        code = (row.get("code") or "").strip()
        if code not in PRODUCT_ORDER:
            continue

        # Hatalı kayıt trend hesabına katılmaz.
        if row.get("status") == "ERROR":
            continue

        pct = _to_float(row.get("spread_pct"))
        if pct is None:
            continue

        bucket = runs.setdefault(
            run_at,
            {
                "dt": run_dt,
                "USD": None,
                "EUR": None,
                "XAU": None,
            },
        )

        current = bucket[code]
        if current is None or pct < current:
            bucket[code] = pct

    result = []
    for item in runs.values():
        result.append(
            (
                item["dt"],
                {
                    "USD": item["USD"],
                    "EUR": item["EUR"],
                    "XAU": item["XAU"],
                },
            )
        )

    result.sort(key=lambda item: item[0])
    return result


def _build_current_sheet(
    wb: Workbook,
    latest_run_at: str | None,
    latest_rows: list[dict],
) -> None:
    # İLK SEKME: Excel açıldığında burası görünecek.
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
            run_dt.time().replace(tzinfo=None)
            if run_dt
            else "",
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

        for code, (
            buy_col,
            sell_col,
            spread_col,
            pct_col,
        ) in layout.items():
            item = row_map.get(code)

            if not item:
                continue

            buy = _to_float(item.get("buy"))
            sell = _to_float(item.get("sell"))

            if buy is not None:
                ws.cell(excel_row, buy_col, buy)

            if sell is not None:
                ws.cell(excel_row, sell_col, sell)

            buy_letter = ws.cell(
                1,
                buy_col,
            ).column_letter
            sell_letter = ws.cell(
                1,
                sell_col,
            ).column_letter
            spread_letter = ws.cell(
                1,
                spread_col,
            ).column_letter

            ws.cell(
                excel_row,
                spread_col,
                (
                    f'=IF(OR('
                    f'{buy_letter}{excel_row}="",'
                    f'{sell_letter}{excel_row}=""),'
                    f'"",'
                    f'{sell_letter}{excel_row}-'
                    f'{buy_letter}{excel_row})'
                ),
            )

            ws.cell(
                excel_row,
                pct_col,
                (
                    f'=IFERROR('
                    f'{spread_letter}{excel_row}/'
                    f'{buy_letter}{excel_row},'
                    f'"")'
                ),
            )

        ws.cell(excel_row, 1).number_format = "dd.mm.yyyy"
        ws.cell(excel_row, 2).number_format = "hh:mm:ss"
        ws.cell(excel_row, 1).alignment = Alignment(
            horizontal="center"
        )
        ws.cell(excel_row, 2).alignment = Alignment(
            horizontal="center"
        )

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

        # Dolar / Euro / Gram Altın için en düşük makas %
        # yeşil olarak vurgulanır.
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


def _build_history_sheet(
    wb: Workbook,
    history: list[dict],
) -> None:
    # İKİNCİ SEKME.
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

    # İstenen sıralama:
    # Banka/Sağlayıcı -> Ürün -> Tarih/Saat
    sorted_history = sorted(
        history,
        key=_history_sort_key,
    )

    for excel_row, item in enumerate(
        sorted_history,
        start=2,
    ):
        run_dt = _parse_dt(item.get("run_at"))
        scraped_dt = _parse_dt(
            item.get("scraped_at")
        )

        buy = _to_float(item.get("buy"))
        sell = _to_float(item.get("sell"))
        site_spread = _to_float(
            item.get("site_spread")
        )
        site_spread_pct = _to_float(
            item.get("site_spread_pct")
        )

        raw_status = item.get("status", "")

        ws.cell(
            excel_row,
            1,
            run_dt.date() if run_dt else "",
        )
        ws.cell(
            excel_row,
            2,
            run_dt.time().replace(tzinfo=None)
            if run_dt
            else "",
        )
        provider_name = item.get("provider", "")
        provider_cell = ws.cell(
            excel_row,
            3,
            provider_name,
        )
        provider_cell.fill = _provider_fill(
            provider_name
        )
        provider_cell.font = Font(bold=True)

        ws.cell(
            excel_row,
            4,
            item.get("product", ""),
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
            (
                f'=IF(OR('
                f'E{excel_row}="",'
                f'F{excel_row}=""),'
                f'"",'
                f'F{excel_row}-E{excel_row})'
            ),
        )
        ws.cell(
            excel_row,
            8,
            (
                f'=IFERROR('
                f'G{excel_row}/E{excel_row},'
                f'"")'
            ),
        )

        # CSV'de teknik durum kodu korunur,
        # Excel'de anlaşılır Türkçe gösterilir.
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
            (
                site_spread_pct / 100.0
                if site_spread_pct is not None
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

        source_cell = ws.cell(
            excel_row,
            10,
        )
        if source_cell.value:
            source_cell.hyperlink = source_cell.value
            source_cell.style = "Hyperlink"

        status_cell = ws.cell(
            excel_row,
            9,
        )
        status_cell.fill = _status_fill(
            str(raw_status)
        )
        status_cell.font = Font(bold=True)
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


def _build_summary_sheet(
    wb: Workbook,
    latest_run_at: str | None,
    latest_rows: list[dict],
    history: list[dict],
) -> None:
    # ÜÇÜNCÜ SEKME.
    ws = wb.create_sheet("OZET")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    ws["A1"] = "Döviz ve Altın Kur Takip Özeti"
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

    run_dt = _parse_dt(latest_run_at)

    providers = {
        row.get("provider")
        for row in latest_rows
        if row.get("provider")
    }

    error_count = sum(
        row.get("status") == "ERROR"
        for row in latest_rows
    )
    control_count = sum(
        row.get("status") == "KONTROL"
        for row in latest_rows
    )

    labels = [
        ("A3", "Son Çekim Tarihi"),
        ("A4", "Son Çekim Saati"),
        ("A5", "Toplam Sağlayıcı"),
        ("A6", "Son Çekim Toplam Kayıt"),
        ("A7", "HATA"),
        ("A8", "KONTROL GEREKLİ"),
    ]

    for coord, label in labels:
        ws[coord] = label
        ws[coord].font = Font(bold=True)
        ws[coord].fill = PatternFill(
            "solid",
            fgColor="EAF2F8",
        )
        ws[coord].border = THIN_BORDER

    ws["B3"] = (
        run_dt.date()
        if run_dt
        else ""
    )
    ws["B4"] = (
        run_dt.time().replace(tzinfo=None)
        if run_dt
        else ""
    )
    ws["B5"] = len(providers)
    ws["B6"] = len(latest_rows)
    ws["B7"] = error_count
    ws["B8"] = control_count

    ws["B3"].number_format = "dd.mm.yyyy"
    ws["B4"].number_format = "hh:mm:ss"

    for row in range(3, 9):
        ws[f"B{row}"].border = THIN_BORDER
        ws[f"B{row}"].alignment = Alignment(
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
            row=11,
            column=col,
            value=header,
        )

    _style_header(
        ws,
        11,
        1,
        len(headers),
    )

    code_names = {
        "USD": "DOLAR",
        "EUR": "EURO",
        "XAU": "GRAM ALTIN",
    }

    row_no = 12

    for code in ("USD", "EUR", "XAU"):
        product_rows = [
            row
            for row in latest_rows
            if row.get("code") == code
        ]

        valid_rows = []

        for row in product_rows:
            buy = _to_float(row.get("buy"))
            sell = _to_float(row.get("sell"))
            pct = _to_float(
                row.get("spread_pct")
            )

            if (
                row.get("status") != "ERROR"
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
                key=lambda item: item[0],
            )
            if valid_rows
            else None
        )

        ws.cell(
            row=row_no,
            column=1,
            value=code_names[code],
        )
        ws.cell(
            row=row_no,
            column=2,
            value=len(product_rows),
        )

        if best:
            pct, row, buy, sell = best

            # CSV spread_pct yüzde biriminde tutuluyor.
            ws.cell(
                row=row_no,
                column=3,
                value=pct / 100.0,
            )
            ws.cell(
                row=row_no,
                column=4,
                value=row.get("provider"),
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
            ).number_format = "0.00%"
            ws.cell(
                row=row_no,
                column=5,
            ).number_format = "#,##0.0000"
            ws.cell(
                row=row_no,
                column=6,
            ).number_format = "#,##0.0000"

        for col in range(1, 7):
            ws.cell(
                row=row_no,
                column=col,
            ).border = THIN_BORDER

        row_no += 1

    # -------------------------------------------------
    # TREND GRAFİĞİ
    # Ayrı TREND sekmesi yok.
    #
    # ÖNEMLİ:
    # Excel varsayılan olarak gizli satır/sütunlardaki verileri
    # grafikte göstermeyebilir. Bu yüzden grafik verisini gizli
    # X:AA kolonlarına koymuyoruz.
    #
    # Yardımcı grafik verisi OZET sayfasında grafiğin altında,
    # A39:D... aralığında tutulur. Böylece Excel'de grafik
    # kesin olarak görünür.
    # -------------------------------------------------
    trend = _build_trend_data(history)

    helper_start_row = 39
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

    for helper_row, (
        trend_dt,
        values,
    ) in enumerate(
        trend,
        start=helper_start_row + 1,
    ):
        ws.cell(
            helper_row,
            1,
            trend_dt.strftime(
                "%d.%m %H:%M"
            ),
        )

        for col, code in enumerate(
            ("USD", "EUR", "XAU"),
            start=2,
        ):
            value = values.get(code)
            cell = ws.cell(
                helper_row,
                col,
            )

            if value is not None:
                cell.value = value / 100.0
                cell.number_format = "0.000%"

    if len(trend) >= 2:
        chart = LineChart()
        chart.style = 10
        chart.title = (
            "Zamana Göre En Düşük Makas % Değişimi"
        )
        chart.y_axis.title = (
            "En Düşük Makas %"
        )
        chart.x_axis.title = (
            "Çekim Zamanı"
        )
        chart.height = 8.5
        chart.width = 17
        chart.legend.position = "b"

        helper_end_row = (
            helper_start_row + len(trend)
        )

        data = Reference(
            ws,
            min_col=2,
            max_col=4,
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
        chart.set_categories(cats)

        try:
            chart.y_axis.numFmt = "0.000%"
        except Exception:
            pass

        ws.add_chart(
            chart,
            "A17",
        )

    _set_widths(
        ws,
        {
            "A": 22,
            "B": 18,
            "C": 18,
            "D": 24,
            "E": 16,
            "F": 16,
        },
    )


def build_excel(
    history_path: str | Path,
    output_path: str | Path,
) -> None:
    """
    Tek bir kalıcı Excel üretir.

    Geçmiş veri CSV'de birikir. Excel her çalışmada
    bu CSV'nin TAMAMINDAN tekrar oluşturulur; böylece
    eski günler kaybolmaz ve output/banka_kurlari.xlsx
    aynı dosya olarak güncellenmeye devam eder.
    """
    history = read_history(history_path)

    if not history:
        raise RuntimeError(
            "Excel oluşturmak için geçmiş veri bulunamadı."
        )

    latest_run_at, latest_rows = _latest_run_rows(
        history
    )

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    # İstenen sekme sırası:
    # 1) GUNCEL_KURLAR
    # 2) GECMIS
    # 3) OZET
    _build_current_sheet(
        wb,
        latest_run_at,
        latest_rows,
    )
    _build_history_sheet(
        wb,
        history,
    )
    _build_summary_sheet(
        wb,
        latest_run_at,
        latest_rows,
        history,
    )

    # Excel açıldığında GUNCEL_KURLAR açılsın.
    wb.active = 0

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output_path = Path(output_path)
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Her seferinde AYNI dosya yolu güncellenir.
    wb.save(output_path)
