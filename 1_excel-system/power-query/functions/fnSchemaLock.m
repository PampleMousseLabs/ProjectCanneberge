(tbl as table) as table =>
let
    // =========================================================
    // REQUIRED FINAL SCHEMA (your system contract)
    // =========================================================
    AllowedColumns =
        {
            "Ticker",
            "Line Item",
            "TTM",
            "Current",
            "2025",
            "2024",
            "2023",
            "2022",
            "2021",
            "Key"
        },

    ExistingCols = Table.ColumnNames(tbl),

    // =========================================================
    // STEP 1: REMOVE UNWANTED COLUMNS
    // =========================================================
    RemovedExtras =
        Table.SelectColumns(
            tbl,
            List.Intersect({ExistingCols, AllowedColumns}),
            MissingField.Ignore
        ),

    // =========================================================
    // STEP 2: ENSURE ALL EXPECTED COLUMNS EXIST
    // =========================================================
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

    // =========================================================
    // STEP 3: ORDER COLUMNS CONSISTENTLY
    // =========================================================
    Reordered =
        Table.ReorderColumns(AddMissing, AllowedColumns, MissingField.Ignore)

in
    Reordered