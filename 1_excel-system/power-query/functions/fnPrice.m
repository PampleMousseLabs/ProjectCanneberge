shared fnPrice = let
    fnPrice = (ticker as text) as nullable number =>
    let
        Json =
            try Json.Document(
                Web.Contents(
                    "https://finnhub.io/api/v1/quote",
                    [
                        Query = [
                            symbol = Text.Upper(ticker),
                            token = "PUT_YOUR_API_KEY_HERE"
                        ]
                    ]
                )
            )
            otherwise null,

        Price =
            if Json = null then
                null
            else
                try Json[c]
                otherwise null
    in
        Price
in
    fnPrice;
