shared Companies = let
    Source = Excel.CurrentWorkbook(){[Name="tblIngest"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Ticker", type text}})
in
    #"Changed Type";
