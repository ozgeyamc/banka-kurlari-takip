from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from io import StringIO
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config.settings import (
    REQUEST_TIMEOUT_SECONDS,
    SPREAD_PCT_TOLERANCE,
    SPREAD_TOLERANCE,
    USER_AGENT,
)
from core.parsing import clean_text, normalize_header, parse_tr_decimal
from core.validation import calculate_spread, validate_record


def _session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        status=3,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
            "Cache-Control": "no-cache",
        }
    )
    return session


def fetch_html(url: str) -> str:
    response = _session().get(url, timeout=REQUEST_TIMEOUT_SECONDS)
    response.raise_for_status()

    if len(response.text) < 1000:
        raise RuntimeError(
            f"Beklenenden kısa HTML döndü ({len(response.text)} karakter): {url}"
        )

    return response.text


def _header_map(headers: list[str]) -> dict[str, int]:
    mapping = {}

    for i, header in enumerate(headers):
        norm = normalize_header(header)

        if "banka" in norm or "kurum" in norm:
            mapping.setdefault("provider", i)
        elif norm == "alis" or "alis" in norm:
            mapping.setdefault("buy", i)
        elif norm == "satis" or "satis" in norm:
            mapping.setdefault("sell", i)
        elif "makas" in norm and ("yuzde" in norm or "%" in clean_text(header)):
            mapping.setdefault("site_spread_pct", i)
        elif "makas" in norm:
            mapping.setdefault("site_spread", i)

    return mapping


def _table_is_target(headers: list[str]) -> bool:
    normalized = [normalize_header(h) for h in headers]
    joined = " | ".join(normalized)

    return (
        ("banka" in joined or "kurum" in joined)
        and "alis" in joined
        and "satis" in joined
        and "makas" in joined
    )


def _extract_with_bs4(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    candidates = []

    for table in soup.find_all("table"):
        header_cells = table.find_all("th")
        headers = [clean_text(x.get_text(" ", strip=True)) for x in header_cells]

        if not headers:
            first_row = table.find("tr")
            if first_row:
                headers = [
                    clean_text(x.get_text(" ", strip=True))
                    for x in first_row.find_all(["td", "th"])
                ]

        if not headers or not _table_is_target(headers):
            continue

        mapping = _header_map(headers)
        if not {"provider", "buy", "sell"}.issubset(mapping):
            continue

        rows = []

        for tr in table.find_all("tr"):
            cells = tr.find_all("td")
            if not cells:
                continue

            values = [clean_text(td.get_text(" ", strip=True)) for td in cells]

            if len(values) <= max(mapping.values()):
                continue

            provider = values[mapping["provider"]]
            if not provider or normalize_header(provider) in {"banka", "kurum"}:
                continue

            rows.append(
                {
                    "provider": provider,
                    "buy_raw": values[mapping["buy"]],
                    "sell_raw": values[mapping["sell"]],
                    "site_spread_raw": (
                        values[mapping["site_spread"]]
                        if "site_spread" in mapping
                        and mapping["site_spread"] < len(values)
                        else None
                    ),
                    "site_spread_pct_raw": (
                        values[mapping["site_spread_pct"]]
                        if "site_spread_pct" in mapping
                        and mapping["site_spread_pct"] < len(values)
                        else None
                    ),
                }
            )

        if rows:
            candidates.append(rows)

    return max(candidates, key=len) if candidates else []


def _extract_with_pandas(html: str) -> list[dict]:
    try:
        tables = pd.read_html(StringIO(html))
    except ValueError:
        return []

    candidates = []

    for df in tables:
        headers = [clean_text(c) for c in df.columns]

        if not _table_is_target(headers):
            continue

        mapping = _header_map(headers)
        if not {"provider", "buy", "sell"}.issubset(mapping):
            continue

        rows = []

        for _, row in df.iterrows():
            values = [clean_text(v) for v in row.tolist()]

            if len(values) <= max(mapping.values()):
                continue

            provider = values[mapping["provider"]]

            if not provider or provider.lower() == "nan":
                continue

            rows.append(
                {
                    "provider": provider,
                    "buy_raw": values[mapping["buy"]],
                    "sell_raw": values[mapping["sell"]],
                    "site_spread_raw": (
                        values[mapping["site_spread"]]
                        if "site_spread" in mapping
                        and mapping["site_spread"] < len(values)
                        else None
                    ),
                    "site_spread_pct_raw": (
                        values[mapping["site_spread_pct"]]
                        if "site_spread_pct" in mapping
                        and mapping["site_spread_pct"] < len(values)
                        else None
                    ),
                }
            )

        if rows:
            candidates.append(rows)

    return max(candidates, key=len) if candidates else []


def extract_rate_rows(html: str) -> list[dict]:
    rows = _extract_with_bs4(html)
    if rows:
        return rows

    rows = _extract_with_pandas(html)
    if rows:
        return rows

    raise RuntimeError(
        "Banka/sağlayıcı kuru tablosu bulunamadı. "
        "Doviz.com sayfa yapısı değişmiş olabilir."
    )


def scrape_product(code: str, product: str, url: str) -> list[dict]:
    html = fetch_html(url)
    raw_rows = extract_rate_rows(html)

    if not raw_rows:
        raise RuntimeError(f"Tablo bulundu ancak hiç veri satırı alınamadı: {url}")

    scraped_at = datetime.now(ZoneInfo("Europe/Istanbul")).isoformat(timespec="seconds")
    result = []

    for row in raw_rows:
        buy = parse_tr_decimal(row.get("buy_raw"))
        sell = parse_tr_decimal(row.get("sell_raw"))
        site_spread = parse_tr_decimal(row.get("site_spread_raw"))
        site_spread_pct = parse_tr_decimal(row.get("site_spread_pct_raw"))

        spread, spread_pct = calculate_spread(buy, sell)

        status, note = validate_record(
            buy=buy,
            sell=sell,
            spread=spread,
            spread_pct=spread_pct,
            site_spread=site_spread,
            site_spread_pct=site_spread_pct,
            spread_tolerance=Decimal(str(SPREAD_TOLERANCE)),
            spread_pct_tolerance=Decimal(str(SPREAD_PCT_TOLERANCE)),
        )

        result.append(
            {
                "scraped_at": scraped_at,
                "product": product,
                "code": code,
                "provider": row["provider"],
                "buy": buy,
                "sell": sell,
                "spread": spread,
                "spread_pct": spread_pct,
                "site_spread": site_spread,
                "site_spread_pct": site_spread_pct,
                "source_url": url,
                "status": status,
                "note": note,
            }
        )

    return result
