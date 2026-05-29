shared fnEngine = (tkr as text, ftype as text) as table =>
let
    Result =
        if ftype = "IS" then fnIS(tkr)
        else if ftype = "BS" then fnBS(tkr)
        else if ftype = "CFS" then fnCFS(tkr)
        else if ftype = "Ratio" then fnRatio(tkr)
        else if ftype = "Beta" then fnBeta(tkr)
        else null,

    Output =
        if Result = null then null
        else Table.AddColumn(Result, "Ticker", each Text.Lower(tkr))
in
    Output;
