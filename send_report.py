from __future__ import annotations

import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo


EXCEL_PATH = Path("output/banka_kurlari.xlsx")


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Eksik GitHub Secret / environment variable: {name}"
        )
    return value


def send_excel_report() -> None:
    """
    Oluşturulan banka_kurlari.xlsx dosyasını
    TEK EK olarak belirtilen adrese gönderir.

    Varsayılan SMTP ayarları Gmail içindir.
    Gmail kullanıyorsanız MAIL_APP_PASSWORD alanına
    normal hesap şifresi değil Google App Password girin.
    """
    mail_user = _required_env("MAIL_USER")
    mail_password = _required_env("MAIL_APP_PASSWORD")
    mail_to = _required_env("MAIL_TO")

    smtp_host = (
        os.getenv("SMTP_HOST", "").strip()
        or "smtp.gmail.com"
    )
    smtp_port = int(
        os.getenv("SMTP_PORT", "").strip()
        or "465"
    )

    if not EXCEL_PATH.exists():
        raise FileNotFoundError(
            f"Excel bulunamadı: {EXCEL_PATH}"
        )

    now_tr = datetime.now(
        ZoneInfo("Europe/Istanbul")
    )

    message = EmailMessage()
    message["From"] = mail_user
    message["To"] = mail_to
    message["Subject"] = (
        "Döviz ve Altın Kur Takip - "
        + now_tr.strftime("%d.%m.%Y")
    )

    message.set_content(
        "Merhaba,\n\n"
        "Günlük Dolar, Euro ve Gram Altın kur takip Excel'i "
        "ekte yer almaktadır.\n\n"
        f"Çalışma zamanı: {now_tr.strftime('%d.%m.%Y %H:%M:%S')} "
        "(Türkiye saati)\n\n"
        "Bu e-postaya yalnızca banka_kurlari.xlsx dosyası eklenmiştir."
    )

    excel_bytes = EXCEL_PATH.read_bytes()

    message.add_attachment(
        excel_bytes,
        maintype="application",
        subtype=(
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        filename="banka_kurlari.xlsx",
    )

    with smtplib.SMTP_SSL(
        smtp_host,
        smtp_port,
        timeout=60,
    ) as smtp:
        smtp.login(
            mail_user,
            mail_password,
        )
        smtp.send_message(message)

    print(
        f"[MAIL] Excel gönderildi -> {mail_to}"
    )


if __name__ == "__main__":
    send_excel_report()
