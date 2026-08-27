from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from zoneinfo import ZoneInfo

from bs4 import BeautifulSoup
from playwright.sync_api import (
    Browser,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

from config.settings import (
    MIN_ROWS_PER_PRODUCT,
    NAVIGATION_TIMEOUT_MS,
    SPREAD_PCT_TOLERANCE,
    SPREAD_TOLERANCE,
    TABLE_TIMEOUT_MS,
    USER_AGENT,
)

from core.parsing import (
    clean_text,
    normalize_header,
    parse_tr_decimal,
)

from core.validation import (
    calculate_spread,
    validate_record,
)


def _header_map(
    headers: list[str],
) -> dict[str, int]:

    mapping: dict[str, int] = {}

    for index, header in enumerate(headers):

        normalized = normalize_header(
            header
        )

        if (
            "banka" in normalized
            or "kurum" in normalized
            or "saglayici" in normalized
        ):
            mapping.setdefault(
                "provider",
                index,
            )

        elif "alis" in normalized:
            mapping.setdefault(
                "buy",
                index,
            )

        elif "satis" in normalized:
            mapping.setdefault(
                "sell",
                index,
            )

        elif (
            "makas" in normalized
            and "yuzde" in normalized
        ):
            mapping.setdefault(
                "site_spread_pct",
                index,
            )

        elif "makas" in normalized:
            mapping.setdefault(
                "site_spread",
                index,
            )

    return mapping


def _is_target_table(
    headers: list[str],
) -> bool:

    normalized = " | ".join(
        normalize_header(item)
        for item in headers
    )

    return (
        (
            "banka" in normalized
            or "kurum" in normalized
            or "saglayici" in normalized
        )
        and "alis" in normalized
        and "satis" in normalized
        and "makas" in normalized
    )


def extract_rate_rows(
    html: str,
) -> list[dict]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    candidates: list[list[dict]] = []

    for table in soup.find_all("table"):

        first_row = table.find("tr")

        if first_row is None:
            continue

        header_cells = (
            first_row.find_all(
                ["th", "td"]
            )
        )

        headers = [
            clean_text(
                cell.get_text(
                    " ",
                    strip=True,
                )
            )
            for cell in header_cells
        ]

        if (
            not headers
            or not _is_target_table(
                headers
            )
        ):
            continue

        mapping = _header_map(
            headers
        )

        if not {
            "provider",
            "buy",
            "sell",
        }.issubset(mapping):

            continue

        rows: list[dict] = []

        for tr in table.find_all(
            "tr"
        )[1:]:

            cells = tr.find_all(
                ["td", "th"]
            )

            values = [
                clean_text(
                    cell.get_text(
                        " ",
                        strip=True,
                    )
                )
                for cell in cells
            ]

            if (
                not values
                or len(values)
                <= max(mapping.values())
            ):
                continue

            provider = values[
                mapping["provider"]
            ]

            if not provider:
                continue

            rows.append(
                {
                    "provider":
                        provider,

                    "buy_raw":
                        values[
                            mapping["buy"]
                        ],

                    "sell_raw":
                        values[
                            mapping["sell"]
                        ],

                    "site_spread_raw":
                        (
                            values[
                                mapping[
                                    "site_spread"
                                ]
                            ]
                            if (
                                "site_spread"
                                in mapping
                                and mapping[
                                    "site_spread"
                                ] < len(values)
                            )
                            else None
                        ),

                    "site_spread_pct_raw":
                        (
                            values[
                                mapping[
                                    "site_spread_pct"
                                ]
                            ]
                            if (
                                "site_spread_pct"
                                in mapping
                                and mapping[
                                    "site_spread_pct"
                                ] < len(values)
                            )
                            else None
                        ),
                }
            )

        if rows:
            candidates.append(
                rows
            )

    if not candidates:
        raise RuntimeError(
            "Banka/sağlayıcı tablosu "
            "bulunamadı. "
            "Sayfa yapısı değişmiş "
            "veya içerik yüklenmemiş olabilir."
        )

    # Sayfada birden fazla tablo varsa
    # en fazla satıra sahip olan
    # banka/sağlayıcı tablosunu al.
    return max(
        candidates,
        key=len,
    )


def _new_context(
    browser: Browser,
):

    return browser.new_context(
        locale="tr-TR",
        timezone_id="Europe/Istanbul",

        viewport={
            "width": 1440,
            "height": 1100,
        },

        user_agent=USER_AGENT,

        extra_http_headers={
            "Accept-Language":
                "tr-TR,tr;q=0.9,"
                "en-US;q=0.8,en;q=0.7",

            "Cache-Control":
                "no-cache",
        },
    )


def _load_html(
    browser: Browser,
    url: str,
) -> str:

    context = _new_context(
        browser
    )

    page = context.new_page()

    page.set_default_timeout(
        TABLE_TIMEOUT_MS
    )

    try:

        response = page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=NAVIGATION_TIMEOUT_MS,
        )

        if (
            response is not None
            and response.status >= 400
        ):
            raise RuntimeError(
                f"HTTP "
                f"{response.status}: "
                f"{url}"
            )

        # Sayfadaki dinamik verilerin
        # yüklenmesini beklemeyi dene.
        try:
            page.wait_for_load_state(
                "networkidle",
                timeout=10_000,
            )

        except PlaywrightTimeoutError:
            pass

        # En az bir tablo oluşmasını bekle.
        try:
            (
                page.locator("table")
                .first
                .wait_for(
                    state="attached",
                    timeout=TABLE_TIMEOUT_MS,
                )
            )

        except PlaywrightTimeoutError:
            pass

        # Son JS güncellemeleri için
        # kısa bekleme.
        page.wait_for_timeout(
            2_000
        )

        return page.content()

    finally:
        context.close()


def scrape_all_products(
    products: dict,
) -> tuple[
    list[dict],
    list[dict],
]:

    all_rows: list[dict] = []
    failures: list[dict] = []

    with sync_playwright() as playwright:

        browser = (
            playwright.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
        )

        try:

            for code, cfg in (
                products.items()
            ):

                url = cfg["url"]
                product = cfg["product"]

                print(
                    f"[{code}] "
                    f"{url}"
                )

                try:

                    html = _load_html(
                        browser,
                        url,
                    )

                    raw_rows = (
                        extract_rate_rows(
                            html
                        )
                    )

                    if (
                        len(raw_rows)
                        < MIN_ROWS_PER_PRODUCT
                    ):
                        raise RuntimeError(
                            "Şüpheli derecede "
                            "az satır bulundu: "
                            f"{len(raw_rows)} "
                            f"(minimum "
                            f"{MIN_ROWS_PER_PRODUCT})"
                        )

                    scraped_at = (
                        datetime.now(
                            ZoneInfo(
                                "Europe/Istanbul"
                            )
                        )
                        .isoformat(
                            timespec="seconds"
                        )
                    )

                    seen = set()

                    product_rows = []

                    for raw in raw_rows:

                        provider = (
                            clean_text(
                                raw[
                                    "provider"
                                ]
                            )
                        )

                        if not provider:
                            continue

                        # Aynı sağlayıcı DOM'da
                        # yanlışlıkla iki defa
                        # görünürse tekrar yazma.
                        provider_key = (
                            provider.casefold()
                        )

                        if (
                            provider_key
                            in seen
                        ):
                            continue

                        seen.add(
                            provider_key
                        )

                        buy = (
                            parse_tr_decimal(
                                raw.get(
                                    "buy_raw"
                                )
                            )
                        )

                        sell = (
                            parse_tr_decimal(
                                raw.get(
                                    "sell_raw"
                                )
                            )
                        )

                        site_spread = (
                            parse_tr_decimal(
                                raw.get(
                                    "site_spread_raw"
                                )
                            )
                        )

                        site_spread_pct = (
                            parse_tr_decimal(
                                raw.get(
                                    "site_spread_pct_raw"
                                )
                            )
                        )

                        (
                            spread,
                            spread_pct,
                        ) = calculate_spread(
                            buy,
                            sell,
                        )

                        (
                            status,
                            note,
                        ) = validate_record(
                            buy=buy,
                            sell=sell,

                            spread=spread,
                            spread_pct=spread_pct,

                            site_spread=
                                site_spread,

                            site_spread_pct=
                                site_spread_pct,

                            spread_tolerance=
                                Decimal(
                                    SPREAD_TOLERANCE
                                ),

                            spread_pct_tolerance=
                                Decimal(
                                    SPREAD_PCT_TOLERANCE
                                ),
                        )

                        product_rows.append(
                            {
                                "scraped_at":
                                    scraped_at,

                                "product":
                                    product,

                                "code":
                                    code,

                                "provider":
                                    provider,

                                "buy":
                                    buy,

                                "sell":
                                    sell,

                                "spread":
                                    spread,

                                "spread_pct":
                                    spread_pct,

                                "site_spread":
                                    site_spread,

                                "site_spread_pct":
                                    site_spread_pct,

                                "source_url":
                                    url,

                                "status":
                                    status,

                                "note":
                                    note,
                            }
                        )

                    if (
                        len(product_rows)
                        < MIN_ROWS_PER_PRODUCT
                    ):
                        raise RuntimeError(
                            "Temizlik sonrası "
                            "şüpheli derecede az "
                            "sağlayıcı kaldı: "
                            f"{len(product_rows)}"
                        )

                    all_rows.extend(
                        product_rows
                    )

                    print(
                        f"[{code}] "
                        f"{len(product_rows)} "
                        "sağlayıcı bulundu."
                    )

                except Exception as exc:

                    failures.append(
                        {
                            "code":
                                code,

                            "product":
                                product,

                            "url":
                                url,

                            "error":
                                str(exc),
                        }
                    )

        finally:
            browser.close()

    return (
        all_rows,
        failures,
    )
