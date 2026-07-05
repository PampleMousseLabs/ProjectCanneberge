(ticker as text) as table =>
let
    // REFACTORED: Read year anchors from Control sheet
    NFY = Date.Year(Excel.CurrentWorkbook(){[Name="NextFiscalYear"]}[Content]{0}[Column1]),
    NFY1 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_1"]}[Content]{0}[Column1]),
    NFY2 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_2"]}[Content]{0}[Column1]),

    NFY_txt = Text.From(NFY),
    NFY1_txt = Text.From(NFY1),
    NFY2_txt = Text.From(NFY2),

    Source = Excel.CurrentWorkbook(){[Name="tblForwardEst_Raw"]}[Content],

    CleanTicker = Text.Lower(Text.Trim(ticker)),

    Filtered = Table.SelectRows(
        Source,
        each Text.Lower(Text.Trim([Ticker])) = CleanTicker
    ),

    // REFACTORED: Dynamic column typing instead of hardcoded years
    Typed = Table.TransformColumnTypes(
        Filtered,
        {
            {NFY_txt, type number},
            {NFY1_txt, type number},
            {NFY2_txt, type number}
        }
    ),

    // REFACTORED: Dynamic column selection instead of hardcoded years
    Ordered = Table.SelectColumns(
        Typed,
        {"Key", "Ticker", "Line Item", NFY_txt, NFY1_txt, NFY2_txt},
        MissingField.Ignore
    )
in
    Ordered
