shared fnIS = (ticker as text) as table =>
let
    Url =
        "https://stockanalysis.com/stocks/"
        & Text.Lower(Text.Trim(ticker))
        & "/financials/",

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

    Renamed =
        Table.RenameColumns(
            AddTicker,
            List.Transform(
                Table.ColumnNames(AddTicker),
                (col) =>
                    if Text.StartsWith(col, "TTM") then {col, "TTM"}
                    else if Text.Contains(col, "FY 2025") then {col, "2025"}
                    else if Text.Contains(col, "FY 2024") then {col, "2024"}
                    else if Text.Contains(col, "FY 2023") then {col, "2023"}
                    else if Text.Contains(col, "FY 2022") then {col, "2022"}
                    else if Text.Contains(col, "FY 2021") then {col, "2021"}
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
    fnSchemaLock(AddKey);
