# banka-kurlari-takip

Doviz.com üzerindeki **Dolar (USD)**, **Euro (EUR)** ve **Gram Altın (XAU/GRAM_ALTIN)** sayfalarında yer alan
tüm sağlayıcıların alış/satış değerlerini toplayan başlangıç projesi.

## v0.1 hedefi

Bu sürüm özellikle veri çekme katmanını doğrular.

- Kurum/sağlayıcı whitelist'i yoktur.
- Sayfadaki tabloda bulunan tüm satırlar alınır.
- Alış ve satıştan makas tekrar hesaplanır.
- Sitedeki makas ve makas yüzdesi ile çapraz kontrol yapılır.
- Eksik veya mantıksız değerler `ERROR` / `KONTROL` olarak işaretlenir.
- İlk test çıktısı `data/latest_rates.csv` dosyasına yazılır.
- Henüz Excel geçmişi ve GitHub Actions eklenmemiştir; bunlar v0.2 adımıdır.

## Kaynaklar

- USD: https://kur.doviz.com/serbest-piyasa/amerikan-dolari
- EUR: https://kur.doviz.com/serbest-piyasa/euro
- Gram Altın: https://altin.doviz.com/gram-altin

## Kurulum

```bash
python -m pip install -r requirements.txt
python main.py
```

Test:

```bash
python -m unittest discover -s tests -v
```

## Beklenen çıktı

`data/latest_rates.csv`

Kolonlar:

- scraped_at
- product
- code
- provider
- buy
- sell
- spread
- spread_pct
- site_spread
- site_spread_pct
- source_url
- status
- note

## Tasarım kararı

Sağlayıcı isimleri elle tanımlanmaz. Yeni bir kurum/kanal Doviz.com tablosuna eklenirse,
HTML tablosunda yer aldığı sürece scraper otomatik olarak almaya çalışır.

Veri bulunamadığında sıfır yazılmaz.
