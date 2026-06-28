let
    // =========================================================
    // 1. TICKERS
    // =========================================================
    Source =
        Excel.CurrentWorkbook(){[Name="tblIngest"]}[Content],

    CleanTickers =
        Table.SelectRows(
            Source,
            each [Ticker] <> null and Text.Trim([Ticker]) <> ""
        ),

    TickerList =
        List.Buffer(
            List.Transform(
                CleanTickers[Ticker],
                each Text.Lower(Text.Trim(_))
            )
        ),

    // =========================================================
    // 2. INVOKE fnForwardEst PER TICKER
    // =========================================================
    RawResults =
        List.Transform(
            TickerList,
            (t) => try Table.Buffer(fnForwardEst(t)) otherwise null
        ),

    // =========================================================
    // 3. FILTER VALID TABLES
    // =========================================================
    ValidTables =
        List.Select(
            RawResults,
            each _ <> null and Value.Is(_, type table)
        ),

    // =========================================================
    // 4. COMBINE
    // =========================================================
    Combined =
        if List.Count(ValidTables) = 0 then
            #table({}, {})
        else
            Table.Combine(ValidTables)
in
    Combined