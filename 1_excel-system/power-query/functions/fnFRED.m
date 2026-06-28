(seriesID as text) as record =>
let
    apiKey = Excel.CurrentWorkbook(){[Name="KeyFRED"]}[Content]{0}[Column1],
    url = "https://api.stlouisfed.org/fred/series/observations?series_id=" 
          & seriesID 
          & "&api_key=" & apiKey 
          & "&sort_order=desc&limit=1&file_type=json",
    raw = Web.Contents(url, [Headers=[Accept="application/json"]]),
    json = Json.Document(raw),
    obs = json[observations],
    latest = obs{0},
    dateVal = latest[date],
    valueVal = latest[value],
    result = [
        SeriesID = seriesID,
        AsOfDate = dateVal,
        LatestValue = valueVal
    ]
in
    result