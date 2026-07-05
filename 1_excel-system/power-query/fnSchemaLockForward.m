(tbl as table) as table =>
let
    NFY = Date.Year(Excel.CurrentWorkbook(){[Name="NextFiscalYear"]}[Content]{0}[Column1]),
    NFY1 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_1"]}[Content]{0}[Column1]),
    NFY2 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_2"]}[Content]{0}[Column1]),

    // =========================================================
    // FORWARD SCHEMA — for fnForwardEst
    // =========================================================
    AllowedColumns =
        {
            "Ticker",
            "Line Item",
            Text.From(NFY),
            Text.From(NFY1),
            Text.From(NFY2),
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
