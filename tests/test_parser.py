import unittest

from decimal import Decimal

from core.parsing import (
    normalize_header,
    parse_tr_decimal,
)

from scrapers.doviz_com import (
    extract_rate_rows,
)


SAMPLE_HTML = """
<html>
<body>

<table>
  <thead>
    <tr>
      <th>Banka</th>
      <th>Alış</th>
      <th>Satış</th>
      <th>Makas</th>
      <th>Makas(%)</th>
    </tr>
  </thead>

  <tbody>

    <tr>
      <td>Kapalıçarşı</td>
      <td>48,0700</td>
      <td>48,0800</td>
      <td>0,0100</td>
      <td>%0,02</td>
    </tr>

    <tr>
      <td>Akbank</td>
      <td>47,3390</td>
      <td>48,6890</td>
      <td>1,3500</td>
      <td>%2,85</td>
    </tr>

    <tr>
      <td>Örnek Sağlayıcı</td>
      <td>7.119,25</td>
      <td>7.125,57</td>
      <td>6,32</td>
      <td>%0,09</td>
    </tr>

  </tbody>
</table>

</body>
</html>
"""


class ParserTests(
    unittest.TestCase
):

    def test_turkish_headers(
        self
    ):

        self.assertEqual(
            normalize_header(
                "Alış"
            ),
            "alis",
        )

        self.assertEqual(
            normalize_header(
                "Satış"
            ),
            "satis",
        )

        self.assertEqual(
            normalize_header(
                "Makas(%)"
            ),
            "makas yuzde",
        )

    def test_turkish_decimals(
        self
    ):

        self.assertEqual(
            parse_tr_decimal(
                "48,0700"
            ),
            Decimal(
                "48.0700"
            ),
        )

        self.assertEqual(
            parse_tr_decimal(
                "7.119,25"
            ),
            Decimal(
                "7119.25"
            ),
        )

        self.assertEqual(
            parse_tr_decimal(
                "%2,39"
            ),
            Decimal(
                "2.39"
            ),
        )

        self.assertIsNone(
            parse_tr_decimal(
                "-"
            )
        )

    def test_all_providers_are_kept(
        self
    ):

        rows = extract_rate_rows(
            SAMPLE_HTML
        )

        names = [
            row["provider"]
            for row in rows
        ]

        self.assertEqual(
            names,
            [
                "Kapalıçarşı",
                "Akbank",
                "Örnek Sağlayıcı",
            ],
        )


if __name__ == "__main__":
    unittest.main()
