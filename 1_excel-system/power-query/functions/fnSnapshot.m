shared fnSnapshot = (ticker as text) as record =>
let
    // =========================================================
    // ONE WEB CALL PER TICKER (CORE OPTIMIZATION)
    // =========================================================
    Url =
        "https://stockanalysis.com/stocks/"
        & Text.Lower(Text.Trim(ticker))
        & "/financials/",

    Raw = Web.Contents(Url),
    Source = Web.Page(Raw),
    Tables = Source[Data],

    // =========================================================
    // SNAPSHOT STRUCTURE (cached per ticker execution)
    // =========================================================
    Snapshot =
        [
            IS     = try Tables{0} otherwise null,
            BS     = try Tables{1} otherwise null,
            CFS    = try Tables{2} otherwise null,
            Ratio  = try Tables{3} otherwise null
        ]
in
    Snapshot;
