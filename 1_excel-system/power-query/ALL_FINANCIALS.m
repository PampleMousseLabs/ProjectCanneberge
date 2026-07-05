let
    // REFACTORED: Read year anchors from Control sheet
    LFY = Date.Year(Excel.CurrentWorkbook(){[Name="FiscalYearEnd"]}[Content]{0}[Column1]),
    NFY = Date.Year(Excel.CurrentWorkbook(){[Name="NextFiscalYear"]}[Content]{0}[Column1]),
    NFY1 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_1"]}[Content]{0}[Column1]),
    NFY2 = Date.Year(Excel.CurrentWorkbook(){[Name="NFY_2"]}[Content]{0}[Column1]),

    // REFACTORED: Dynamic column list
    OutputColumns =
        {
            "Key",
            "Ticker",
            "Line Item",
            "TTM",
            "Current",
            Text.From(LFY - 4),
            Text.From(LFY - 3),
            Text.From(LFY - 2),
            Text.From(LFY - 1),
            Text.From(LFY),
            Text.From(NFY),
            Text.From(NFY1),
            Text.From(NFY2)
        },

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
    // 4. CLEAN FINAL OUTPUT (REFACTORED: dynamic column list)
    // =========================================================
    Cleaned =
        Table.SelectColumns(
            CombinedTable,
            List.Intersect(
                {
                    OutputColumns,
                    Table.ColumnNames(CombinedTable)
                }
            ),
            MissingField.Ignore
        )
in
    Cleaned
