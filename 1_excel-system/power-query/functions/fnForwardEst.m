(ticker as text, ms_slug as text) as table =>
let
    // =========================================================
    // 1. STOCKANALYSIS — REVENUE & EPS ESTIMATES
    // =========================================================
    SA_Url =
        "https://stockanalysis.com/stocks/"
        & Text.Lower(Text.Trim(ticker))
        & "/forecast/",

    SA_Source = Web.Page(Web.Contents(SA_Url)),
    SA_Tables = SA_Source[Data],

    ToNumber = (v as any) as any =>
        if v = null then null
        else if v = "Pro" then null
        else if v = "-" then null
        else if v = "" then null
        else
            let
                s = Text.Trim(Text.From(v)),
                suffix = Text.End(s, 1),
                base = Text.Start(s, Text.Length(s) - 1),
                result =
                    if suffix = "B" then Number.From(base) * 1000
                    else if suffix = "M" then Number.From(base)
                    else if suffix = "K" then Number.From(base) / 1000
                    else Number.From(s)
            in
                try result otherwise null,

    ProcessSATable = (tbl as table, prefix as text) as table =>
        let
            FirstCol = Table.ColumnNames(tbl){0},
            Labeled = Table.RenameColumns(tbl, {{FirstCol, "Line Item"}}),
            Renamed =
                Table.RenameColumns(
                    Labeled,
                    List.Transform(
                        Table.ColumnNames(Labeled),
                        (col) =>
                            if Text.Contains(col, "2026") then {col, "2026"}
                            else if Text.Contains(col, "2027") then {col, "2027"}
                            else if Text.Contains(col, "2028") then {col, "2028"}
                            else if Text.Contains(col, "2029") then {col, "2029"}
                            else if Text.Contains(col, "2030") then {col, "2030"}
                            else {col, col}
                    )
                ),
            Prefixed =
                Table.TransformColumns(
                    Renamed,
                    {{"Line Item", each prefix & " " & Text.Lower(Text.Trim(Text.From(_))), type text}}
                ),
            YearCols = {"2026", "2027", "2028", "2029", "2030"},
            ExistingYearCols = List.Intersect({Table.ColumnNames(Prefixed), YearCols}),
            Converted =
                Table.TransformColumns(
                    Prefixed,
                    List.Transform(ExistingYearCols, (col) => {col, each ToNumber(_), type any})
                ),
            WithTicker =
                Table.AddColumn(Converted, "Ticker", each Text.Lower(Text.Trim(ticker))),
            WithKey =
                Table.AddColumn(WithTicker, "Key", each [Ticker] & "|" & [Line Item])
        in
            WithKey,

    RevTable = try SA_Tables{4} otherwise null,
    EpsTable = try SA_Tables{6} otherwise null,

    RevClean =
        if RevTable = null then #table({}, {})
        else ProcessSATable(RevTable, "rev est"),

    EpsClean =
        if EpsTable = null then #table({}, {})
        else ProcessSATable(EpsTable, "eps est"),

    SA_Combined = Table.Combine({RevClean, EpsClean}),

    // =========================================================
    // 2. MARKETSCREENER — EBITDA & EBIT ESTIMATES
    // =========================================================
    MS_Url =
        "https://www.marketscreener.com/quote/stock/"
        & Text.Trim(ms_slug)
        & "/finances/",

    MS_Source = Web.Page(Web.Contents(MS_Url)),
    MS_Tables = MS_Source[Data],
    MS_RawTable = try MS_Tables{2} otherwise null,

    ProcessMSTable = (tbl as table) as table =>
        let
            FirstCol = Table.ColumnNames(tbl){0},

            // Keep only rows we want — Net sales, EBITDA, EBIT, Net income
            KeepRows = {"net sales", "ebitda", "ebit", "net income"},

            Labeled = Table.RenameColumns(tbl, {{FirstCol, "Line Item"}}),

            // Normalize Line Item to lowercase for filtering
            Lowered =
                Table.TransformColumns(
                    Labeled,
                    {{"Line Item", each Text.Lower(Text.Trim(Text.From(_))), type text}}
                ),

            // Filter to only keep rows we want
            Filtered =
                Table.SelectRows(
                    Lowered,
                    each List.Contains(KeepRows, [Line Item])
                ),

            // Rename year columns
            Renamed =
                Table.RenameColumns(
                    Filtered,
                    List.Transform(
                        Table.ColumnNames(Filtered),
                        (col) =>
                            if Text.Contains(col, "2026") then {col, "2026"}
                            else if Text.Contains(col, "2027") then {col, "2027"}
                            else if Text.Contains(col, "2028") then {col, "2028"}
                            else if Text.Contains(col, "2029") then {col, "2029"}
                            else if Text.Contains(col, "2030") then {col, "2030"}
                            else if Text.Contains(col, "2025") then {col, "2025"}
                            else if Text.Contains(col, "2024") then {col, "2024"}
                            else if Text.Contains(col, "2023") then {col, "2023"}
                            else {col, col}
                    )
                ),

            // Prefix line items to match naming convention
            Prefixed =
                Table.TransformColumns(
                    Renamed,
                    {{
                        "Line Item",
                        each
                            if _ = "net sales" then "ms rev est"
                            else if _ = "ebitda" then "ms ebitda est"
                            else if _ = "ebit" then "ms ebit est"
                            else if _ = "net income" then "ms net income est"
                            else _,
                        type text
                    }}
                ),

            // Convert all year columns to numbers
            YearCols = {"2023","2024","2025","2026","2027","2028","2029","2030"},
            ExistingYearCols =
                List.Intersect({Table.ColumnNames(Prefixed), YearCols}),

            Converted =
                Table.TransformColumns(
                    Prefixed,
                    List.Transform(
                        ExistingYearCols,
                        (col) => {col, each try Number.From(Text.Replace(Text.From(_), ",", "")) otherwise null, type any}
                    )
                ),

            WithTicker =
                Table.AddColumn(Converted, "Ticker", each Text.Lower(Text.Trim(ticker))),

            WithKey =
                Table.AddColumn(WithTicker, "Key", each [Ticker] & "|" & [Line Item])
        in
            WithKey,

    MS_Clean =
        if MS_RawTable = null then #table({}, {})
        else ProcessMSTable(MS_RawTable),

    // =========================================================
    // 3. COMBINE SA + MS
    // =========================================================
    Combined = Table.Combine({SA_Combined, MS_Clean}),

    // =========================================================
    // 4. COLUMN ORDER
    // =========================================================
    Ordered =
        Table.SelectColumns(
            Combined,
            List.Intersect(
                {
                    {
                        "Key",
                        "Ticker",
                        "Line Item",
                        "2023",
                        "2024",
                        "2025",
                        "2026",
                        "2027",
                        "2028",
                        "2029",
                        "2030"
                    },
                    Table.ColumnNames(Combined)
                }
            ),
            MissingField.Ignore
        )
in
    Ordered