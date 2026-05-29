shared Financials_Diagnostics = let
    // =========================================================
    // 1. Tickers
    // =========================================================
    Tickers =
        List.Distinct(Companies[Ticker]),

   CleanTickers =
    List.RemoveNulls(
        List.Transform(
            Tickers,
            each
                if _ = null then null
                else Text.Lower(Text.Trim(Text.From(_)))
        )
    ),

    // =========================================================
    // 2. TIMING WRAPPER (VBA FRIENDLY OUTPUT)
    // =========================================================
    TimeFn =
        (fn as function, tkr as text, functionName as text) as record =>
        let
            Start = DateTime.LocalNow(),

            Result =
                try fn(tkr) otherwise null,

            End = DateTime.LocalNow(),

            Rows =
                try if Result = null then 0 else Table.RowCount(Result) otherwise 0,

            Status =
                if Result = null then "FAIL"
                else "OK"
        in
            [
                Ticker = tkr,
                Function = functionName,
                Status = Status,
                Rows = Rows,
                StartTime = Start,
                EndTime = End,
                DurationMs =
                    Number.From(End - Start) * 24 * 60 * 60 * 1000
            ],

    // =========================================================
    // 3. RUN ALL FUNCTIONS
    // =========================================================
    IS =
        List.Transform(CleanTickers, each TimeFn(fnIS, _, "IS")),

    BS =
        List.Transform(CleanTickers, each TimeFn(fnBS, _, "BS")),

    CFS =
        List.Transform(CleanTickers, each TimeFn(fnCFS, _, "CFS")),

    Ratio =
        List.Transform(CleanTickers, each TimeFn(fnRatio, _, "Ratio")),

    Beta =
        List.Transform(CleanTickers, each TimeFn(fnBeta, _, "Beta")),

    // =========================================================
    // 4. FLATTEN INTO SINGLE TABLE (VBA SAFE)
    // =========================================================
    Combined =
        Table.FromRecords(
            List.Combine({IS, BS, CFS, Ratio, Beta})
        )
in
    Combined;
