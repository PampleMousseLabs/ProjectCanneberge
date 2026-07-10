// =========================================================
// fnPriceHistory
// Pulls daily OHLCV + AdjClose from Yahoo Finance JSON API
// for a single ticker over an explicit date window.
//
// Inputs:
//   ticker     (text)  e.g. "AMZN" or "^SPX"
//   startDate  (date)  first date of window (inclusive-ish)
//   endDate    (date)  last date of window
//
// Output: table with columns Date, Open, High, Low, Close, AdjClose, Volume
//         sorted ascending (oldest first)
//
// Notes:
//   - Uses Yahoo's undocumented but stable JSON chart endpoint
//   - Endpoint requires unix timestamps (seconds since 1970-01-01 UTC)
//   - Start/end dates are supplied by caller (typically from named ranges
//     on Inputs sheet that already include any desired buffer days)
//   - For indices (^SPX etc), AdjClose == Close (no dividends)
//   - For stocks, AdjClose is split/dividend-adjusted — use this for returns
// =========================================================
(ticker as text, startDate as date, endDate as date) as table =>
let
    // Convert dates to unix seconds
    epoch = #date(1970, 1, 1),
    startUnix = Number.From(startDate - epoch) * 86400,
    endUnix   = Number.From(endDate - epoch) * 86400,
    
    // Build URL
    Url =
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        & ticker
        & "?period1=" & Number.ToText(startUnix)
        & "&period2=" & Number.ToText(endUnix)
        & "&interval=1d",
    
    // Pull JSON
    Source = Json.Document(Web.Contents(Url)),
    Result = Source[chart][result]{0},
    
    // Extract parallel arrays
    Timestamps    = Result[timestamp],
    Quote         = Result[indicators][quote]{0},
    AdjCloseList  = Result[indicators][adjclose]{0}[adjclose],
    Opens         = Quote[open],
    Highs         = Quote[high],
    Lows          = Quote[low],
    Closes        = Quote[close],
    Volumes       = Quote[volume],
    
    // Zip arrays into rows
    RowCount = List.Count(Timestamps),
    Rows = List.Transform(
        {0..RowCount - 1},
        (i) => {
            Date.From(#datetime(1970,1,1,0,0,0) + #duration(0,0,0,Timestamps{i})),
            Opens{i},
            Highs{i},
            Lows{i},
            Closes{i},
            AdjCloseList{i},
            Volumes{i}
        }
    ),
    
    // Assemble typed table
    Table = #table(
        type table [
            Date     = date,
            Open     = nullable number,
            High     = nullable number,
            Low      = nullable number,
            Close    = nullable number,
            AdjClose = nullable number,
            Volume   = nullable Int64.Type
        ],
        Rows
    ),
    
    Sorted = Table.Sort(Table, {{"Date", Order.Ascending}})
in
    Sorted