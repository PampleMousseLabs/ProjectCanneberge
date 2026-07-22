from Canneberge.Sources.stockanalysis import StockAnalysisClient


class SourceDataService:
    """
    Coordinates source-data pulls.

    Batch 1 only wires StockAnalysis.
    MarketScreener, FRED, and Beta/Vol will be added after this works.
    """

    def __init__(self, project_inputs, progress_callback=None):
        self.project_inputs = project_inputs
        self.progress = progress_callback or (lambda message: None)

    def refresh_stockanalysis(self):
        tickers = self.project_inputs.active_public_tickers
        statements = ["IS", "BS", "CFS", "Ratios"]

        client = StockAnalysisClient()
        all_results = {stmt: [] for stmt in statements}

        for idx, ticker in enumerate(tickers):
            self.progress(f"StockAnalysis: {ticker} ({idx + 1}/{len(tickers)})")

            for stmt in statements:
                try:
                    df = client.fetch_statement(ticker, stmt)

                    if df is not None and not df.empty:
                        records = df.to_dict("records")
                        all_results[stmt].extend(records)
                        self.progress(f"  {ticker} {stmt}: {len(records)} rows")
                    else:
                        self.progress(f"  {ticker} {stmt}: No data")

                except Exception as e:
                    self.progress(f"  {ticker} {stmt}: Error - {str(e)}")

        return all_results