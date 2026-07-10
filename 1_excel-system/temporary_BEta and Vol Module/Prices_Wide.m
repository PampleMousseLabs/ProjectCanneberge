// =========================================================
// Prices_Wide
// Combines Index_Prices + ALL_TickerPrices into a wide-format
// AdjClose matrix: one Date column + one column per ticker.
//
// Column order: Date, IndexTicker, then tickers in tblTickers order
// Values: AdjClose (nulls where ticker had no data on that date)
// =========================================================
let
    // Load both source tables (AdjClose only)
    IndexRaw = Index_Prices,
    IndexTrimmed = Table.SelectColumns(IndexRaw, {"Ticker","Date","AdjClose"}),
    
    TickersRaw = ALL_TickerPrices,
    TickersTrimmed = Table.SelectColumns(TickersRaw, {"Ticker","Date","AdjClose"}),
    
    // Union into one long table
    Combined = Table.Combine({IndexTrimmed, TickersTrimmed}),
    
    // Pivot to wide format
    Pivoted = Table.Pivot(
        Combined,
        List.Distinct(Combined[Ticker]),
        "Ticker",
        "AdjClose",
        List.Sum
    ),
    
    // Sort by date ascending
    Sorted = Table.Sort(Pivoted, {{"Date", Order.Ascending}}),
    
    // Reorder: Date first, Index second, then tickers in tblTickers order
    IndexTicker = Excel.CurrentWorkbook(){[Name="SelectedIndexTicker"]}[Content]{0}[Column1],
    TickerOrder = Excel.CurrentWorkbook(){[Name="tblTickers"]}[Content][Ticker],
    DesiredOrder = List.Combine({{"Date", IndexTicker}, TickerOrder}),
    
    // Filter DesiredOrder to only columns that actually exist (in case a ticker failed)
    ActualCols = Table.ColumnNames(Sorted),
    FinalOrder = List.Select(DesiredOrder, each List.Contains(ActualCols, _)),
    
    Reordered = Table.ReorderColumns(Sorted, FinalOrder)
in
    Reordered