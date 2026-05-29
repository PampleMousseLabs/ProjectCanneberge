let
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
                    try Table.Buffer(fnIS(t)) otherwise null,
                    try Table.Buffer(fnBS(t)) otherwise null,
                    try Table.Buffer(fnCFS(t)) otherwise null,
                    try Table.Buffer(fnRatio(t)) otherwise null,
                    try Table.Buffer(fnBeta(t)) otherwise null,
                    try Table.Buffer(fnForwardEst(t)) otherwise null
                })
            )
        ),

    // =========================================================
    // 3. COMBINE TABLES
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
                    {
                        "Key",
                        "Ticker",
                        "Line Item",
                        "TTM",
                        "Current",
                        "2021",
                        "2022",
                        "2023",
                        "2024",
                        "2025",
                        "2026",
                        "2027",
                        "2028",
                        "2029",
                        "2030"
                    },
                    Table.ColumnNames(CombinedTable)
                }
            ),
            MissingField.Ignore
        )
in
    Cleaned
