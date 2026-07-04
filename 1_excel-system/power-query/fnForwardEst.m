(ticker as text) as table =>
let
    Source = Excel.CurrentWorkbook(){[Name="tblForwardEst_Raw"]}[Content],

    CleanTicker = Text.Lower(Text.Trim(ticker)),

    Filtered = Table.SelectRows(
        Source,
        each Text.Lower(Text.Trim([Ticker])) = CleanTicker
    ),

    Typed = Table.TransformColumnTypes(
        Filtered,
        {
            {"2026", type number},
            {"2027", type number},
            {"2028", type number}
        }
    ),

    Ordered = Table.SelectColumns(
        Typed,
        {"Key", "Ticker", "Line Item", "2026", "2027", "2028"},
        MissingField.Ignore
    )
in
    Ordered
