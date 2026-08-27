from __future__ import annotations

from decimal import Decimal


def calculate_spread(
    buy: Decimal | None,
    sell: Decimal | None,
):
    if buy is None or sell is None or buy <= 0:
        return None, None

    spread = sell - buy

    spread_pct = (
        spread / buy
    ) * Decimal("100")

    return spread, spread_pct


def validate_record(
    buy,
    sell,
    spread,
    spread_pct,
    site_spread,
    site_spread_pct,
    spread_tolerance: Decimal,
    spread_pct_tolerance: Decimal,
):
    errors = []
    warnings = []

    if buy is None:
        errors.append(
            "Alış değeri bulunamadı"
        )

    if sell is None:
        errors.append(
            "Satış değeri bulunamadı"
        )

    if buy is not None and buy <= 0:
        errors.append(
            "Alış değeri pozitif değil"
        )

    if sell is not None and sell <= 0:
        errors.append(
            "Satış değeri pozitif değil"
        )

    if (
        buy is not None
        and sell is not None
        and sell < buy
    ):
        errors.append(
            "Satış alıştan küçük"
        )

    if errors:
        return (
            "ERROR",
            "; ".join(errors),
        )

    # Bizim hesapladığımız makas ile
    # Doviz.com makasını karşılaştır.
    if (
        site_spread is not None
        and spread is not None
    ):
        if (
            abs(site_spread - spread)
            > spread_tolerance
        ):
            warnings.append(
                f"Makas farkı: "
                f"site={site_spread}, "
                f"hesap={spread}"
            )

    if (
        site_spread_pct is not None
        and spread_pct is not None
    ):
        if (
            abs(
                site_spread_pct
                - spread_pct
            )
            > spread_pct_tolerance
        ):
            warnings.append(
                f"Makas % farkı: "
                f"site={site_spread_pct}, "
                f"hesap={spread_pct}"
            )

    if warnings:
        return (
            "KONTROL",
            "; ".join(warnings),
        )

    return "OK", ""
