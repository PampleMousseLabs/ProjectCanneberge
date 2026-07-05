(ticker as text) as table =>
let
    // REFACTORED: Read year anchors from Control sheet
    LFY = Date.Year(Excel.CurrentWorkbook(){[Name="FiscalYearEnd"]}[Content]{0}[Column1]),
    LFY_1 = LFY - 1,
    LFY_2 = LFY - 2,
    LFY_3 = LFY - 3,
    LFY_4 = LFY - 4,

    Url =
        "https://stockanalysis.com/stocks/"
        & Text.Lower(Text.Trim(ticker))
        & "/financials/cash-flow-statement/",

    Source = Web.Page(Web.Contents(Url)),
    Tables = Source[Data],

    RawTable =
        List.First(
            List.Select(Tables, each Table.RowCount(_) > 5 and Table.ColumnCount(_) > 2),
            null
        ),

    Clean =
        if RawTable = null then
            #table({}, {})
        else
            fnCleanFinancialTable(RawTable),

    FirstCol = Table.ColumnNames(Clean){0},

    Standardized =
        Table.RenameColumns(
            Clean,
            {
                {FirstCol, "Line Item"}
            },
            MissingField.Ignore
        ),

    AddTicker =
        Table.AddColumn(Standardized, "Ticker", each Text.Lower(Text.Trim(ticker))),

    // REFACTORED: Dynamic year mapping instead of hardcoded years
    Renamed =
        Table.RenameColumns(
            AddTicker,
            List.Transform(
                Table.ColumnNames(AddTicker),
                (col) =>
                    if Text.StartsWith(col, "TTM") then {col, "TTM"}
                    else if Text.Contains(col, "FY " & Text.From(LFY)) then {col, Text.From(LFY)}
                    else if Text.Contains(col, "FY " & Text.From(LFY_1)) then {col, Text.From(LFY_1)}
                    else if Text.Contains(col, "FY " & Text.From(LFY_2)) then {col, Text.From(LFY_2)}
                    else if Text.Contains(col, "FY " & Text.From(LFY_3)) then {col, Text.From(LFY_3)}
                    else if Text.Contains(col, "FY " & Text.From(LFY_4)) then {col, Text.From(LFY_4)}
                    else if Text.StartsWith(col, "Current") then {col, "Current"}
                    else {col, col}
            )
        ),

    AddKey =
        Table.AddColumn(
            Renamed,
            "Key",
            each [Ticker] & "|" & Text.Lower([Line Item])
        )
in
    fnSchemaLockFinancials(AddKey)
