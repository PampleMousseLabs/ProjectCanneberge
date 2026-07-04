let
    // =========================================================
    // 1. Load tickers
    // =========================================================
    Source = Companies,

    CleanTickers =
        Table.SelectRows(
            Source,
            each [Ticker] <> null and Text.Trim([Ticker]) <> ""
        ),

    TickerList =
        List.Buffer(
            List.Distinct(
                List.Transform(CleanTickers[Ticker], each Text.Trim(_))
            )
        ),

    // =========================================================
    // 2. Invoke fnCFS safely per ticker
    // =========================================================
    RawResults =
        List.Transform(
            TickerList,
            (t) =>
                let
                    result =
                        try fnRatio(t) otherwise null
                in
                    result
        ),

    // =========================================================
    // 3. Remove invalid outputs (critical stability fix)
    // =========================================================
    ValidTables =
        List.Select(
            RawResults,
            each _ <> null and Value.Is(_, type table)
        ),

    // =========================================================
    // 4. Combine safely
    // =========================================================
    Combined =
        if List.Count(ValidTables) = 0 then
            #table({}, {})
        else
            Table.Combine(ValidTables)
in
    Combined
