// =========================================================
// Index_Prices
// Pulls the selected market index using explicit start/end
// dates from Inputs sheet (IndexStartDate includes buffer).
//
// Kept separate from ticker pulls because:
//   - Different window (shorter — only needs BetaHistory)
//   - Downstream beta calc needs it as its own dataset
// =========================================================
let
    IndexTicker = Excel.CurrentWorkbook(){[Name="SelectedIndexTicker"]}[Content]{0}[Column1],
    StartDate   = Date.From(Excel.CurrentWorkbook(){[Name="IndexStartDate"]}[Content]{0}[Column1]),
    EndDate     = Date.From(Excel.CurrentWorkbook(){[Name="ValuationDate"]}[Content]{0}[Column1]),
    
    Source = fnPriceHistory(IndexTicker, StartDate, EndDate),
    Tagged = Table.AddColumn(Source, "Ticker", each IndexTicker, type text),
    Reordered = Table.ReorderColumns(
        Tagged,
        {"Ticker","Date","Open","High","Low","Close","AdjClose","Volume"}
    )
in
    Reordered