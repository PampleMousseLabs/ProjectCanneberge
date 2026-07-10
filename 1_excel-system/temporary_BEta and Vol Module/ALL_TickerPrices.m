// =========================================================
// ALL_TickerPrices
// Invokes fnPriceHistory for each ticker in tblTickers
// using explicit start/end dates from Inputs sheet.
//
// TickerStartDate on Inputs already includes buffer days,
// so no year-rounding gymnastics needed here.
//
// Skips tickers that fail (returns null → filtered out).
// =========================================================
let
    // Load inputs from named ranges
    Tickers   = Excel.CurrentWorkbook(){[Name="tblTickers"]}[Content][Ticker],
    StartDate = Date.From(Excel.CurrentWorkbook(){[Name="TickerStartDate"]}[Content]{0}[Column1]),
    EndDate   = Date.From(Excel.CurrentWorkbook(){[Name="ValuationDate"]}[Content]{0}[Column1]),
    
    // Clean ticker list
    CleanTickers = List.Buffer(
        List.Distinct(
            List.Select(
                List.Transform(Tickers, each Text.Trim(_)),
                each _ <> null and _ <> ""
            )
        )
    ),
    
    // Pull each ticker; wrap in try to prevent single failure killing whole pipeline
    RawResults = List.Transform(
        CleanTickers,
        (t) =>
            let
                pulled = try fnPriceHistory(t, StartDate, EndDate) otherwise null,
                tagged =
                    if pulled = null then null
                    else Table.AddColumn(pulled, "Ticker", each t, type text)
            in
                tagged
    ),
    
    // Drop failed pulls
    ValidTables = List.Select(RawResults, each _ <> null and Value.Is(_, type table)),
    
    // Combine
    Combined =
        if List.Count(ValidTables) = 0 then
            #table(
                type table [Ticker=text, Date=date, Open=nullable number, High=nullable number, Low=nullable number, Close=nullable number, AdjClose=nullable number, Volume=nullable Int64.Type],
                {}
            )
        else
            Table.Combine(ValidTables),
    
    // Put Ticker first
    Reordered = Table.ReorderColumns(
        Combined,
        {"Ticker","Date","Open","High","Low","Close","AdjClose","Volume"}
    )
in
    Reordered