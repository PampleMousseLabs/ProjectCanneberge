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
            List.Transform(
                CleanTickers[Ticker],
                each Text.Trim(_)
            )
        ),

    // =========================================================
    // 2. Invoke fnBeta correctly (THIS IS THE CRITICAL PART)
    // =========================================================
    ResultsList =
        List.Transform(
            TickerList,
            each try fnBeta(_) otherwise null
        ),

    // =========================================================
    // 3. Remove nulls
    // =========================================================
    CleanList =
        List.RemoveNulls(ResultsList),

    // =========================================================
    // 4. Ensure only tables are combined
    // =========================================================
    ValidTables =
        List.Select(
            CleanList,
            each Value.Is(_, type table)
        ),

    // =========================================================
    // 5. Combine
    // =========================================================
    ALL_Beta =
        if List.Count(ValidTables) = 0 then
            #table({"Ticker","LineItem","Value","Key"}, {})
        else
            Table.Combine(ValidTables)

in
    ALL_Beta
