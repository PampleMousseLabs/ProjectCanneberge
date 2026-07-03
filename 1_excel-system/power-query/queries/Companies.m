let
    Source = Excel.CurrentWorkbook(){[Name="tblIngest"]}[Content],
    Typed = Table.SelectRows(
    Table.TransformColumnTypes(Source, {{"Ticker", type text}}),
    each [Ticker] <> null and Text.Trim([Ticker]) <> ""
),
    // Read subject company toggle
    CompanyStatus = 
        try Text.Trim(Excel.CurrentWorkbook(){[Name="CompanyStatus"]}[Content]{0}[Column1])
        otherwise "",

    SubjectTicker =
        try Text.Trim(Excel.CurrentWorkbook(){[Name="SubjectCompanyTicker"]}[Content]{0}[Column1])
        otherwise "",

    // Append subject ticker if publicly traded and not already in tblIngest
    SubjectRow = #table({"Ticker"}, {{SubjectTicker}}),

    AlreadyExists =
        List.Contains(
            List.Transform(Typed[Ticker], each Text.Lower(Text.Trim(_))),
            Text.Lower(SubjectTicker)
        ),

    Combined =
        if Text.Lower(CompanyStatus) = "publicly traded"
            and SubjectTicker <> ""
            and not AlreadyExists
        then Table.Combine({Typed, SubjectRow})
        else Typed
in
    Combined