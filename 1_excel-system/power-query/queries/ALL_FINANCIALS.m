shared ALL_FINANCIALS = let
    // =========================================================
    // 1. TICKERS (BUFFERED)
    // =========================================================
    Tickers =
        List.Buffer(
            List.Transform(
                List.Distinct(Companies[Ticker]),
                each Text.Lower(Text.Trim(_))
            )
        ),

    // =========================================================
    // 2. SINGLE PASS EVALUATION (SAFE TABLE OUTPUTS)
    // =========================================================
    CombinedList =
        List.Combine(
            List.Transform(Tickers, (t) =>
                List.RemoveNulls({
                    try fnIS(t) otherwise null,
                    try fnBS(t) otherwise null,
                    try fnCFS(t) otherwise null,
                    try fnRatio(t) otherwise null,
                    try fnBeta(t) otherwise null
                })
            )
        ),

    // =========================================================
    // 3. COMBINE TABLES (UNCHANGED MODEL)
    // =========================================================
    CombinedTable =
        Table.Combine(CombinedList),

    // =========================================================
    // 4. CLEAN FINAL OUTPUT
    // =========================================================
    Cleaned =
        Table.SelectColumns(
            CombinedTable,
            List.Intersect(
                {
                    Table.ColumnNames(CombinedTable),
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
                    }
                }
            ),
            MissingField.Ignore
        )
in
    Cleaned;
