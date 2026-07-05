(tbl as table) as table =>
let
    // =========================================================
    // BETA SCHEMA — for fnBeta (single value per ticker)
    // =========================================================
    AllowedColumns =
        {
            "Ticker",
            "Line Item",
            "Current",
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
