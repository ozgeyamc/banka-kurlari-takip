from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from config.settings import PRODUCTS
from core.excel_writer import build_excel
from scrapers.doviz_com import scrape_all_products


APP_VERSION = "0.4"

LATEST_PATH = Path("data/latest_rates.csv")
HISTORY_PATH = Path("data/rates_history.csv")
EXCEL_PATH = Path("output/banka_kurlari.xlsx")

FIELDNAMES = [
    "run_at",
    "scraped_at",
    "product",
    "code",
    "provider",
    "buy",
    "sell",
    "spread",
    "spread_pct",
    "site_spread",
    "site_spread_pct",
    "source_url",
    "status",
    "note",
]


def decimal_text(value):
    if value is None:
        return ""
    return format(value, "f")


def serialize_row(row: dict) -> dict:
    output = {
        key: row.get(key, "")
        for key in FIELDNAMES
    }

    for key in (
        "buy",
        "sell",
        "spread",
        "spread_pct",
        "site_spread",
        "site_spread_pct",
    ):
        output[key] = decimal_text(
            row.get(key)
        )

    return output


def write_latest(
    rows: list[dict],
) -> None:
    """
    Sadece son çekimin snapshot'ı.
    Her çalışmada yenilenir.
    """
    LATEST_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with LATEST_PATH.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=FIELDNAMES,
        )
        writer.writeheader()

        for row in rows:
            writer.writerow(
                serialize_row(row)
            )


def append_history(
    rows: list[dict],
) -> None:
    """
    EN ÖNEMLİ KISIM:
    Geçmiş dosyası silinmez/yenilenmez.
    Yeni çekim her çalışmada mevcut CSV'nin
    ALTINA eklenir.

    Excel daha sonra bu dosyanın tamamından
    yeniden oluşturulur.
    """
    HISTORY_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    exists = (
        HISTORY_PATH.exists()
        and HISTORY_PATH.stat().st_size > 0
    )

    with HISTORY_PATH.open(
        "a",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=FIELDNAMES,
        )

        if not exists:
            writer.writeheader()

        for row in rows:
            writer.writerow(
                serialize_row(row)
            )


def main() -> None:
    print(
        f"=== Döviz ve Altın Kur Takip "
        f"v{APP_VERSION} ==="
    )
    print(
        "Kapsam: USD + EUR + GRAM ALTIN"
    )
    print(
        "Sağlayıcı filtresi: YOK\n"
    )

    run_at = datetime.now(
        ZoneInfo("Europe/Istanbul")
    ).isoformat(
        timespec="seconds"
    )

    rows, failures = scrape_all_products(
        PRODUCTS
    )

    # Bir ürün sayfası tamamen hata verdiyse
    # eksik snapshot geçmişe eklenmez.
    if failures:
        print(
            "\n=== SAYFA HATALARI ==="
        )

        for item in failures:
            print(
                f"{item['code']} | "
                f"{item['error']}"
            )

        raise SystemExit(2)

    if not rows:
        raise SystemExit(
            "FATAL: Hiç veri çekilemedi."
        )

    found_codes = {
        row.get("code")
        for row in rows
    }

    missing_codes = (
        set(PRODUCTS)
        - found_codes
    )

    if missing_codes:
        raise SystemExit(
            "FATAL: Şu ürünler tamamen eksik: "
            + ", ".join(
                sorted(missing_codes)
            )
        )

    # Aynı çalıştırmanın bütün satırlarına
    # tek bir run_at atanır.
    for row in rows:
        row["run_at"] = run_at

    # Son snapshot yenilenir.
    write_latest(rows)

    # Geçmişin üzerine yazılmaz;
    # yeni çekim aşağı eklenir.
    append_history(rows)

    # AYNI Excel dosyası geçmiş CSV'nin
    # tamamından yeniden oluşturulur.
    build_excel(
        HISTORY_PATH,
        EXCEL_PATH,
    )

    for code in PRODUCTS:
        product_rows = [
            row
            for row in rows
            if row["code"] == code
        ]

        error_count = sum(
            row["status"] == "ERROR"
            for row in product_rows
        )
        control_count = sum(
            row["status"] == "KONTROL"
            for row in product_rows
        )

        print(
            f"\n[{code}] toplam sağlayıcı: "
            f"{len(product_rows)}"
        )
        print(
            f"[{code}] HATA={error_count} | "
            f"KONTROL GEREKLİ="
            f"{control_count}"
        )

    print(
        f"\nÇekim zamanı (TR): {run_at}"
    )
    print(
        f"Son çekim kayıt sayısı: "
        f"{len(rows)}"
    )
    print(
        f"Güncel CSV: {LATEST_PATH}"
    )
    print(
        f"Geçmiş CSV: {HISTORY_PATH}"
    )
    print(
        f"Tek kalıcı Excel: {EXCEL_PATH}"
    )

    control_rows = [
        row
        for row in rows
        if row["status"] != "OK"
    ]

    if control_rows:
        print(
            "\n=== KONTROL GEREKLİ / "
            "HATA KAYITLARI ==="
        )

        display = {
            "OK": "DOĞRU",
            "KONTROL": "KONTROL GEREKLİ",
            "ERROR": "HATA",
        }

        for row in control_rows:
            print(
                f"{row['code']} | "
                f"{row['provider']} | "
                f"{display.get(row['status'], row['status'])} | "
                f"{row['note']}"
            )


if __name__ == "__main__":
    main()
