PRODUCTS = {
    "USD": {
        "product": "DOLAR",
        "url": "https://kur.doviz.com/serbest-piyasa/amerikan-dolari",
    },
    "EUR": {
        "product": "EURO",
        "url": "https://kur.doviz.com/serbest-piyasa/euro",
    },
    "XAU": {
        "product": "GRAM_ALTIN",
        "url": "https://altin.doviz.com/gram-altin",
    },
}

# Bir ürün sayfasından bundan daha az sağlayıcı gelirse
# scraper başarılı kabul edilmeyecek.
MIN_ROWS_PER_PRODUCT = 10

NAVIGATION_TIMEOUT_MS = 60_000
TABLE_TIMEOUT_MS = 25_000

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/130.0.0.0 Safari/537.36"
)

# Doviz.com değerleri yuvarlayabildiğinden küçük tolerans bırakıyoruz.
SPREAD_TOLERANCE = "0.02"
SPREAD_PCT_TOLERANCE = "0.03"
