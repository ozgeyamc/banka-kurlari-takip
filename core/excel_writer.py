def _style_chart(
    chart,
    title,
    legend=True,
    background=None,
):
    """
    Tüm grafiklerin ortak görünümü.

    ÖNEMLİ:
    - Grid çizgileri yok.
    - Grafikler daha kompakt.
    - Grafik içi veri etiketi yok.
    - Legend altta.
    - Tarihler dik değil, daha okunabilir.
    """

    chart.title = title

    # --------------------------------------------------------
    # BOYUT
    #
    # A / K / U şeklinde yerleştireceğiz.
    # Bu genişlikte grafikler birbirine değmez.
    # --------------------------------------------------------

    chart.width = 18.0
    chart.height = 8.0

    chart.y_axis.title = "Makas %"
    chart.y_axis.numFmt = "0.00%"

    # --------------------------------------------------------
    # ARKA PLAN GRID ÇİZGİLERİ KAPALI
    # --------------------------------------------------------

    chart.y_axis.majorGridlines = None
    chart.x_axis.majorGridlines = None

    chart.y_axis.majorTickMark = "none"
    chart.x_axis.majorTickMark = "none"

    # --------------------------------------------------------
    # EKSENLER
    # --------------------------------------------------------

    try:
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.tickLblSkip = 1
        chart.x_axis.tickMarkSkip = 1
    except Exception:
        pass

    # --------------------------------------------------------
    # GRAFİK DIŞ ALANI
    # --------------------------------------------------------

    try:
        chart.graphical_properties = GraphicalProperties(
            noFill=True
        )
    except Exception:
        pass

    # --------------------------------------------------------
    # PLOT AREA
    #
    # Tek banka grafiklerinde hafif banka rengi.
    # Toplu grafiklerde beyaz.
    # --------------------------------------------------------

    if background:
        try:
            chart.plot_area.graphicalProperties = (
                GraphicalProperties(
                    solidFill=background
                )
            )
        except Exception:
            pass

    # --------------------------------------------------------
    # TARİH YAZILARI
    # --------------------------------------------------------

    try:

        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(
                rot=0
            ),
            p=[
                Paragraph(
                    pPr=ParagraphProperties(
                        defRPr=CharacterProperties(
                            sz=700
                        )
                    ),
                    endParaRPr=CharacterProperties(
                        sz=700
                    ),
                )
            ],
        )

    except Exception:
        pass

    # --------------------------------------------------------
    # LEGEND
    # --------------------------------------------------------

    if legend:

        chart.legend.position = "b"
        chart.legend.overlay = False

    else:

        chart.legend = None


def _build_summary_sheet(
    wb,
    latest_dt,
    latest_rows,
    history,
):

    ws = wb.create_sheet("OZET")

    ws.sheet_view.showGridLines = False

    # ========================================================
    # VERİLER
    # ========================================================

    trends = _build_bank_trends(
        history
    )

    monthly = _monthly_averages(
        history
    )

    (
        data_ws,
        daily_columns,
        monthly_columns,
        monthly_date_col,
        months,
    ) = _build_chart_data_sheet(
        wb,
        trends,
        monthly,
    )

    # ========================================================
    # ANA BAŞLIK
    # ========================================================

    ws.merge_cells("A1:AC1")

    ws["A1"] = (
        "Döviz ve Altın Kur Takip Özeti"
    )

    ws["A1"].fill = TITLE_FILL

    ws["A1"].font = Font(
        bold=True,
        size=16,
        color=DARK_BLUE,
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    ws.row_dimensions[1].height = 28

    # ========================================================
    # ÜST TABLOLAR
    #
    # SOL   : SON VERİ ÇEKİMİ
    # ORTA  : AYLIK ORTALAMA
    # SAĞ   : EN DÜŞÜK MAKAS
    # ========================================================

    _build_latest_table(
        ws,
        latest_dt,
        latest_rows,
    )

    monthly_last_row = (
        _build_monthly_table(
            ws,
            monthly,
            months,
        )
    )

    _build_best_table(
        ws,
        latest_rows,
    )

    # ========================================================
    # TABLOLARIN BİTİŞİ
    # ========================================================

    top_tables_end = max(
        10,
        monthly_last_row,
    )

    # ========================================================
    # GRAFİK YERLEŞİMİ
    #
    # 18 cm genişlikte grafikler için:
    #
    # A
    # K
    # U
    #
    # kullanıyoruz.
    #
    # Böylece kesinlikle birbirlerine binmezler.
    # ========================================================

    CHART_COLUMNS = [
        "A",
        "K",
        "U",
    ]

    # --------------------------------------------------------
    # SATIR ARALIKLARI
    #
    # 8 cm yükseklik yaklaşık 15 satırdan fazla yer kaplıyor.
    # Biz 22 satır bırakıyoruz.
    # --------------------------------------------------------

    ROW_GAP = 22

    DAILY_CHART_ROW = (
        top_tables_end + 4
    )

    MONTHLY_CHART_ROW = (
        DAILY_CHART_ROW
        + ROW_GAP
    )

    INDIVIDUAL_START_ROW = (
        MONTHLY_CHART_ROW
        + ROW_GAP
    )

    # ========================================================
    # GÜNLÜK 5 BANKA TOPLU GRAFİKLER
    # ========================================================

    if trends:

        daily_end_row = (
            len(trends) + 1
        )

        daily_categories = Reference(
            data_ws,
            min_col=1,
            min_row=2,
            max_row=daily_end_row,
        )

        for product_index, code in enumerate(
            (
                "USD",
                "EUR",
                "XAU",
            )
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{PRODUCT_NAMES[code]} "
                    f"- 5 Banka Makas %"
                ),
                legend=True,
            )

            all_values = []

            # ------------------------------------------------
            # 5 BANKA
            # ------------------------------------------------

            for bank in TARGET_BANKS:

                source_col = (
                    daily_columns[code][bank]
                )

                data = Reference(
                    data_ws,
                    min_col=source_col,
                    max_col=source_col,
                    min_row=1,
                    max_row=daily_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                values = []

                for run in trends:

                    value = (
                        run["banks"]
                        .get(bank, {})
                        .get(code)
                    )

                    values.append(value)

                    if value is not None:
                        all_values.append(value)

            # ------------------------------------------------
            # TARİHLER
            # ------------------------------------------------

            chart.set_categories(
                daily_categories
            )

            # ------------------------------------------------
            # BANKA RENKLERİ
            #
            # Çizgi ve nokta aynı renk.
            # ------------------------------------------------

            for index, series in enumerate(
                chart.series
            ):

                if index >= len(TARGET_BANKS):
                    continue

                bank = TARGET_BANKS[index]

                _set_series_color(
                    series,
                    BANK_LINE_COLORS[bank],
                )

            # ------------------------------------------------
            # Y AXIS
            # ------------------------------------------------

            _set_axis_bounds(
                chart,
                all_values,
            )

            # ------------------------------------------------
            # VERİ ETİKETİ YOK
            # ------------------------------------------------

            chart.dLbls = None

            # ------------------------------------------------
            # YERLEŞTİR
            # ------------------------------------------------

            position = (
                f"{CHART_COLUMNS[product_index]}"
                f"{DAILY_CHART_ROW}"
            )

            ws.add_chart(
                chart,
                position,
            )

    # ========================================================
    # AYLIK ORTALAMA GRAFİKLER
    # ========================================================

    if months:

        monthly_end_row = (
            len(months) + 1
        )

        monthly_categories = Reference(
            data_ws,
            min_col=monthly_date_col,
            min_row=2,
            max_row=monthly_end_row,
        )

        for product_index, code in enumerate(
            (
                "USD",
                "EUR",
                "XAU",
            )
        ):

            chart = LineChart()

            _style_chart(
                chart,
                (
                    f"{PRODUCT_NAMES[code]} "
                    f"- Aylık Ortalama Makas %"
                ),
                legend=True,
            )

            all_values = []

            # ------------------------------------------------
            # 5 BANKA
            # ------------------------------------------------

            for bank in TARGET_BANKS:

                source_col = (
                    monthly_columns[code][bank]
                )

                data = Reference(
                    data_ws,
                    min_col=source_col,
                    max_col=source_col,
                    min_row=1,
                    max_row=monthly_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                for month_key in months:

                    value = (
                        monthly
                        .get(month_key, {})
                        .get(bank, {})
                        .get(code)
                    )

                    if value is not None:
                        all_values.append(value)

            chart.set_categories(
                monthly_categories
            )

            # ------------------------------------------------
            # AYLIK:
            #
            # ÇİZGİ VE NOKTA AYNI RENK
            # ------------------------------------------------

            for index, series in enumerate(
                chart.series
            ):

                if index >= len(TARGET_BANKS):
                    continue

                bank = TARGET_BANKS[index]

                _set_series_color(
                    series,
                    BANK_LINE_COLORS[bank],
                )

            _set_axis_bounds(
                chart,
                all_values,
            )

            chart.dLbls = None

            position = (
                f"{CHART_COLUMNS[product_index]}"
                f"{MONTHLY_CHART_ROW}"
            )

            ws.add_chart(
                chart,
                position,
            )

    # ========================================================
    # TEK BANKA GRAFİKLERİ
    #
    # 5 BANKA
    # x
    # DOLAR / EURO / ALTIN
    # ========================================================

    if trends:

        daily_end_row = (
            len(trends) + 1
        )

        daily_categories = Reference(
            data_ws,
            min_col=1,
            min_row=2,
            max_row=daily_end_row,
        )

        for bank_index, bank in enumerate(
            TARGET_BANKS
        ):

            bank_row = (
                INDIVIDUAL_START_ROW
                + bank_index * ROW_GAP
            )

            for product_index, code in enumerate(
                (
                    "USD",
                    "EUR",
                    "XAU",
                )
            ):

                chart = LineChart()

                # --------------------------------------------
                # BANKA ARKA PLAN RENGİ
                # --------------------------------------------

                background = (
                    BANK_CHART_COLORS.get(
                        bank
                    )
                )

                _style_chart(
                    chart,
                    (
                        f"{bank} - "
                        f"{PRODUCT_NAMES[code]} "
                        f"Makas %"
                    ),
                    legend=False,
                    background=background,
                )

                # --------------------------------------------
                # VERİ KOLONU
                # --------------------------------------------

                source_col = (
                    daily_columns[code][bank]
                )

                data = Reference(
                    data_ws,
                    min_col=source_col,
                    max_col=source_col,
                    min_row=1,
                    max_row=daily_end_row,
                )

                chart.add_data(
                    data,
                    titles_from_data=True,
                )

                chart.set_categories(
                    daily_categories
                )

                # --------------------------------------------
                # ÜRÜN RENGİ
                #
                # USD = MAVİ
                # EUR = TURUNCU
                # XAU = ALTIN
                # --------------------------------------------

                product_color = {
                    "USD": USD_COLOR,
                    "EUR": EUR_COLOR,
                    "XAU": XAU_COLOR,
                }[code]

                if chart.series:

                    _set_series_color(
                        chart.series[0],
                        product_color,
                    )

                # --------------------------------------------
                # Y AXIS
                # --------------------------------------------

                values = []

                for run in trends:

                    value = (
                        run["banks"]
                        .get(bank, {})
                        .get(code)
                    )

                    values.append(value)

                _set_axis_bounds(
                    chart,
                    values,
                )

                # --------------------------------------------
                # VERİ ETİKETLERİ KAPALI
                #
                # Böylece grafik içinde yazı yığını oluşmaz.
                # Excel'de noktaya mouse ile gelince değer
                # görülebilir.
                # --------------------------------------------

                chart.dLbls = None

                # --------------------------------------------
                # YERLEŞİM
                # --------------------------------------------

                position = (
                    f"{CHART_COLUMNS[product_index]}"
                    f"{bank_row}"
                )

                ws.add_chart(
                    chart,
                    position,
                )

    # ========================================================
    # OZET SÜTUN GENİŞLİKLERİ
    #
    # Grafiklerin A / K / U başlangıçları için eşit genişlik.
    # ========================================================

    for col in range(
        1,
        31,
    ):

        letter = get_column_letter(
            col
        )

        ws.column_dimensions[
            letter
        ].width = 9.5

    # --------------------------------------------------------
    # TABLOLAR İÇİN ÖZEL GENİŞLİKLER
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15

    ws.column_dimensions["E"].width = 3

    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 16

    ws.column_dimensions["K"].width = 4

    ws.column_dimensions["L"].width = 17
    ws.column_dimensions["M"].width = 13
    ws.column_dimensions["N"].width = 15
    ws.column_dimensions["O"].width = 19
    ws.column_dimensions["P"].width = 15
    ws.column_dimensions["Q"].width = 15

    # ========================================================
    # SATIR YÜKSEKLİKLERİ
    #
    # Excel'de grafik anchor'larının daha stabil kalması için
    # grafik alanlarındaki satırları sabitliyoruz.
    # ========================================================

    final_chart_row = (
        INDIVIDUAL_START_ROW
        + (len(TARGET_BANKS) * ROW_GAP)
    )

    for row in range(
        DAILY_CHART_ROW,
        final_chart_row + 5,
    ):

        ws.row_dimensions[
            row
        ].height = 22

    # ========================================================
    # HELPER SAYFASI
    # ========================================================

    data_ws.sheet_state = "hidden"
