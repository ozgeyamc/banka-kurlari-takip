from __future__ import annotations

from pathlib import Path

import pandas as pd

from config.settings import PRODUCTS
from scrapers.doviz_com import scrape_product


OUTPUT_PATH = Path("data/latest_rates.csv")


def decimal_to_text(value):
    if value is None:
        return None
    return format(value, "f")


def main():
    all_rows = []
    failures = []

    print("=== Doviz.com Kur Takip v0.1 ===")

    for code, cfg in PRODUCTS.items():
        print(f"[{code}] çekiliyor: {cfg['url']}")

        try:
            rows = scrape_product(
                code=code,
                product=cfg["product"],
                url=cfg["url"],
            )
            all_rows.extend(rows)

            error_count = sum(r["status"] == "ERROR" for r in rows)
            control_count = sum(r["status"] == "KONTROL" for r in rows)

            print(
                f"[{code}] {len(rows)} sağlayıcı bulundu | "
                f"ERROR={error_count} | KONTROL={control_count}"
            )

        except Exception as exc:
            failures.append((code, str(exc)))
            print(f"[{code}] HATA: {exc}")

    if not all_rows:
        raise SystemExit("Hiç veri çekilemedi. Çalışma başarısız kabul edildi.")

    df = pd.DataFrame(all_rows)

    numeric_cols = [
        "buy",
        "sell",
        "spread",
        "spread_pct",
        "site_spread",
        "site_spread_pct",
    ]

    for col in numeric_cols:
        df[col] = df[col].map(decimal_to_text)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    print(f"\nÇıktı: {OUTPUT_PATH}")
    print(f"Toplam kayıt: {len(df)}")
    print("\nÜrün bazında kayıt sayıları:")
    print(df.groupby(["code", "product"]).size().to_string())

    print("\nİlk 15 kayıt:")
    cols = ["code", "provider", "buy", "sell", "spread", "spread_pct", "status"]
    print(df[cols].head(15).to_string(index=False))

    if failures:
        print("\nSayfa bazında hatalar:")
        for code, message in failures:
            print(f"- {code}: {message}")
        raise SystemExit(2)


if __name__ == "__main__":
    main()
