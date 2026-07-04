(tbl as table) as table =>
let
    // 1. Standardize "junk text" ? null
    CleanText =
        Table.TransformColumns(
            tbl,
            List.Transform(
                Table.ColumnNames(tbl),
                (col) =>
                    {
                        col,
                        (v) =>
                            if v = null then null
                            else if v = "-" then null
                            else if v = "—" then null
                            else if v = "N/A" then null
                            else if v = "NA" then null
                            else if v = "" then null
                            else v,
                        type any
                    }
            )
        ),

    // 2. Trim text fields safely
    Trimmed =
        Table.TransformColumns(
            CleanText,
            List.Transform(
                Table.ColumnNames(CleanText),
                (col) =>
                    {
                        col,
                        (v) =>
                            if Value.Is(v, type text)
                            then Text.Trim(v)
                            else v,
                        type any
                    }
            )
        )

in
    Trimmed
