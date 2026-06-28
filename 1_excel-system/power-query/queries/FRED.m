let
    refTableRaw = Excel.CurrentWorkbook(){[Name="tblFREDSeries"]}[Content],
    promoted = Table.PromoteHeaders(refTableRaw, [PromoteAllScalars=true]),
    cleaned = Table.SelectColumns(promoted, {"Display Label", "FRED Series ID"}),
    withData = Table.AddColumn(cleaned, "FREDResult", each fnFRED([FRED Series ID])),
    expanded = Table.ExpandRecordColumn(
        withData,
        "FREDResult",
        {"AsOfDate", "LatestValue"},
        {"As Of Date", "Latest Value"}
    ),
    reordered = Table.ReorderColumns(
        expanded,
        {"Display Label", "FRED Series ID", "As Of Date", "Latest Value"}
    ),
    typed = Table.TransformColumnTypes(
        reordered,
        {
            {"Display Label", type text},
            {"FRED Series ID", type text},
            {"As Of Date", type date},
            {"Latest Value", type number}
        }
    )
in
    typed