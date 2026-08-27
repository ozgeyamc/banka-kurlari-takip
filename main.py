from __future__ import annotations

import csv
from pathlib import Path

from config.settings import PRODUCTS

from scrapers.doviz_com import (
    scrape_all_products,
)


OUTPUT_PATH = Path(
    "data/latest_rates.csv"
)


def decimal_text(value):

    if value is None:
        return ""

    return format(
        value,
        "f",
    )


def write_csv(
    rows: list[dict],
) -> None:

    OUTPUT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    fieldnames = [
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

    with OUTPUT_PATH.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:

        writer = csv.DictWriter(
            handle,
            fieldnames=fieldnames,
        )

        writer.writeheader()

        for row in rows:

            output = dict(
                row
            )

            for key in (
                "buy",
                "sell",
                "spread",
                "spread_pct",
                "site_spread",
                "site_spread_pct",
            ):

                output[key] = (
                    decimal_text(
                        output.get(
                            key
                        )
                    )
                )

            writer.writerow(
                output
            )


def main() -> None:

    print(
        "=== Doviz.com "
        "Kur Takip v0.2 ==="
    )

    print(
        "Kapsam: "
        "USD + EUR + GRAM ALTIN"
    )

    print(
        "Sağlayıcı filtresi: YOK\n"
    )

    (
        rows,
        failures,
    ) = scrape_all_products(
        PRODUCTS
    )

    if not rows:
        raise SystemExit(
            "FATAL: "
            "Hiç veri çekilemedi."
        )

    write_csv(
        rows
    )

    for code in PRODUCTS:

        product_rows = [
            row
            for row in rows
            if row["code"] == code
        ]

        print(
            f"\n[{code}] "
            f"toplam sağlayıcı: "
            f"{len(product_rows)}"
        )

        for row in product_rows:

            print(
                "  - "
                f"{row['provider']}"
            )

        error_count = sum(
            row["status"] == "ERROR"
            for row in product_rows
        )

        control_count = sum(
            row["status"] == "KONTROL"
            for row in product_rows
        )

        print(
            f"[{code}] "
            f"ERROR={error_count} | "
            f"KONTROL={control_count}"
        )

    print(
        f"\nToplam kayıt: "
        f"{len(rows)}"
    )

    print(
        f"CSV oluşturuldu: "
        f"{OUTPUT_PATH}"
    )

    error_rows = [
        row
        for row in rows
        if row["status"] == "ERROR"
    ]

    control_rows = [
        row
        for row in rows
        if row["status"] == "KONTROL"
    ]

    if error_rows:

        print(
            "\n=== "
            "ERROR KAYITLARI "
            "==="
        )

        for row in error_rows:

            print(
                f"{row['code']} | "
                f"{row['provider']} | "
                f"{row['note']}"
            )

    if control_rows:

        print(
            "\n=== "
            "KONTROL KAYITLARI "
            "==="
        )

        for row in control_rows:

            print(
                f"{row['code']} | "
                f"{row['provider']} | "
                f"{row['note']}"
            )

    if failures:

        print(
            "\n=== "
            "SAYFA HATALARI "
            "==="
        )

        for item in failures:

            print(
                f"{item['code']} | "
                f"{item['error']}"
            )

        raise SystemExit(2)


if __name__ == "__main__":
    main()
