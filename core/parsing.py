from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re


_TR_MAP = str.maketrans(
    {
        "ı": "i",
        "İ": "i",
        "ş": "s",
        "Ş": "s",
        "ğ": "g",
        "Ğ": "g",
        "ü": "u",
        "Ü": "u",
        "ö": "o",
        "Ö": "o",
        "ç": "c",
        "Ç": "c",
    }
)


def clean_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


def normalize_header(value: object) -> str:
    """
    Örnek:
    Alış -> alis
    Satış -> satis
    Makas(%) -> makas yuzde
    """
    text = clean_text(value).translate(_TR_MAP).lower()
    text = text.replace("%", " yuzde ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def parse_tr_decimal(value: object) -> Decimal | None:
    """
    Türkçe sayı biçimini Decimal'a çevirir.

    48,0700  -> 48.0700
    7.119,25 -> 7119.25
    %2,39    -> 2.39
    """

    text = clean_text(value)

    if not text or text.casefold() in {
        "-",
        "—",
        "–",
        "n/a",
        "nan",
        "none",
    }:
        return None

    text = (
        text.replace("%", "")
        .replace("₺", "")
        .replace("TL", "")
        .replace("tl", "")
        .replace(" ", "")
    )

    # Türkçe biçim:
    # 7.119,25 -> 7119.25
    # 48,0700 -> 48.0700
    if "," in text:
        text = text.replace(".", "").replace(",", ".")

    elif text.count(".") > 1:
        text = text.replace(".", "")

    text = re.sub(r"[^0-9+\-.]", "", text)

    if not text:
        return None

    try:
        return Decimal(text)
    except InvalidOperation:
        return None
