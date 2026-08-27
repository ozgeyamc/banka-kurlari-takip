from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
import unicodedata


def clean_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


def normalize_header(value: object) -> str:
    text = clean_text(value).casefold()
    text = text.translate(str.maketrans({"ı":"i", "ş":"s", "ğ":"g", "ü":"u", "ö":"o", "ç":"c"}))
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("%", " yuzde ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def parse_tr_decimal(value: object) -> Decimal | None:
    """
    Türkçe sayı biçimlerini Decimal'a çevirir.

    Örnek:
      48,0700  -> Decimal('48.0700')
      7.119,25 -> Decimal('7119.25')
      %2,39    -> Decimal('2.39')
      '-'      -> None
    """
    text = clean_text(value)
    if not text or text in {"-", "—", "–", "N/A", "n/a"}:
        return None

    text = text.replace("%", "").replace("₺", "").replace("TL", "").strip()
    text = text.replace(" ", "")

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
