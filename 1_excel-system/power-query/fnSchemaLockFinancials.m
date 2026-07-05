(tbl as table) as table =>
let
    LFY = Date.Year(Excel.CurrentWorkbook(){[Name="FiscalYearEnd"]}[Content]{0}[Column1]),

    // =========================================================
    // FINANCIALS SCHEMA — for fnIS, fnBS, fnCFS
    // =========================================================
    AllowedColumns =
        {
            "Ticker",
            "Line Item",
            "TTM",
            Text.From(LFY),
            Text.From(LFY - 1),
            Text.From(LFY - 2),
            Text.From(LFY - 3),
            Text.From(LFY - 4),
            "Key"
        },

    ExistingCols = Table.ColumnNames(tbl),

    RemovedExtras =
        Table.SelectColumns(
            tbl,
            List.Intersect({ExistingCols, AllowedColumns}),
            MissingField.Ignore
        ),

    AddMissing =
        List.Accumulate(
            AllowedColumns,
            RemovedExtras,
            (state, col) =>
                if List.Contains(Table.ColumnNames(state), col) then
                    state
                else
                    Table.AddColumn(state, col, each null)
        ),

    Reordered =
        Table.ReorderColumns(AddMissing, AllowedColumns, MissingField.Ignore)
in
    Reordered
