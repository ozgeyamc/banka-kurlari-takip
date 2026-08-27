from __future__ import annotations


def calculate_spread(buy, sell):
    if buy is None or sell is None or buy <= 0:
        return None, None

    spread = sell - buy
    spread_pct = (spread / buy) * 100
    return spread, spread_pct


def validate_record(
    buy,
    sell,
    spread,
    spread_pct,
    site_spread,
    site_spread_pct,
    spread_tolerance,
    spread_pct_tolerance,
):
    problems = []
    warnings = []

    if buy is None:
        problems.append("Alış değeri bulunamadı")
    if sell is None:
        problems.append("Satış değeri bulunamadı")

    if buy is not None and buy <= 0:
        problems.append("Alış değeri pozitif değil")
    if sell is not None and sell <= 0:
        problems.append("Satış değeri pozitif değil")

    if buy is not None and sell is not None and sell < buy:
        problems.append("Satış alıştan küçük")

    if problems:
        return "ERROR", "; ".join(problems)

    if site_spread is not None and spread is not None:
        if abs(site_spread - spread) > spread_tolerance:
            warnings.append(
                f"Site makası uyuşmuyor: site={site_spread}, hesap={spread}"
            )

    if site_spread_pct is not None and spread_pct is not None:
        if abs(site_spread_pct - spread_pct) > spread_pct_tolerance:
            warnings.append(
                f"Site makas yüzdesi uyuşmuyor: site={site_spread_pct}, hesap={spread_pct}"
            )

    if warnings:
        return "KONTROL", "; ".join(warnings)

    return "OK", ""
