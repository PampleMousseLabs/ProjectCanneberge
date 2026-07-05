(tkr as text) as table =>
let
    Url =
        "https://stockanalysis.com/stocks/"
        & Text.Lower(Text.Trim(tkr))
        & "/",

    Source = Web.Page(Web.Contents(Url)),
    Tables = Source[Data],

    BetaTable = try Tables{1} otherwise null,

    BetaValue =
        try
            if BetaTable = null then null
            else
                let
                    Clean =
                        Table.TransformColumns(
                            BetaTable,
                            {
                                {"Column1", each Text.Lower(Text.Trim(Text.From(_))), type text},
                                {"Column2", each Text.From(_), type text}
                            }
                        ),

                    Row =
                        Table.SelectRows(Clean, each [Column1] = "beta"),

                    Result =
                        if Table.IsEmpty(Row) then null
                        else Number.From(Row{0}[Column2])
                in
                    Result
        otherwise null,

    Output =
        #table(
            {"Ticker", "Line Item", "Current", "Key"},
            {
                {
                    Text.Lower(Text.Trim(tkr)),
                    "beta",
                    BetaValue,
                    Text.Lower(Text.Trim(tkr)) & "|beta"
                }
            }
        )
in
    fnSchemaLockBeta(Output)
