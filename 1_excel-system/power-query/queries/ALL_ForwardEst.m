let
    // =========================================================
    // 1. TICKERS + MS SLUGS
    // =========================================================
    Source =
        Excel.CurrentWorkbook(){[Name="tblIngest"]}[Content],

    CleanTickers =
        Table.SelectRows(
            Source,
            each [Ticker] <> null and Text.Trim([Ticker]) <> ""
        ),

    SlugSource =
        Excel.CurrentWorkbook(){[Name="MS_Slug"]}[Content],

    TickerSlugList =
        List.Buffer(
            List.Transform(
                List.Positions(CleanTickers[Ticker]),
                (i) => {
                    Text.Trim(CleanTickers[Ticker]{i}),
                    Text.Trim(SlugSource[MS_Slug]{i} ?? "")
                }
            )
        ),

    // =========================================================
    // 2. INVOKE fnForwardEst PER TICKER
    // =========================================================
    RawResults =
        List.Transform(
            TickerSlugList,
            (pair) => try Table.Buffer(fnForwardEst(pair{0}, pair{1})) otherwise null
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