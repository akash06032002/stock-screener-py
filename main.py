import json
import logging
from pathlib import Path

import pandas as pd
import yfinance as yf
import os
from EXCHANGE_FILTER import EXCHANGE
from scipy.stats import skew

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


class StockScreener:
    """
    A class-based refactor that preserves existing functionality:
      - Lookup tickers, filter by exchange
      - Compute metrics (std dev, skewness, mean volume)
      - Check negative daily returns over last N days
      - Update / append metrics to an Excel sheet (preserving other sheets)
      - Filter + rank tickers and write a 'Ranking' sheet
    """

    def __init__(
        self, preferred_exchange: str = EXCHANGE, logger: logging.Logger | None = None
    ):
        self.preferred_exchange = preferred_exchange
        self.log = logger or logging.getLogger(__name__)

    # ------------------------- Lookup & Filter -------------------------

    def lookup_stock(self, name: str) -> dict[str, str] | None:
        """Looks up stock exchange information for a given company name.

        Args:
            name (str): Company name

        Returns:
            str: Dictionary of stock exchanges (key: ticker, value: exchange)
        """
        try:
            lookup_object = yf.Lookup(query=name)
            stock_object = lookup_object.stock.to_json()

            # Convert JSON string to dict
            stock_json = json.loads(stock_object)

            stock_exchange_dict = stock_json.get("exchange")

            return stock_exchange_dict
        except Exception as e:
            self.log.error("Error in lookup_stock, Company name:%s, %s", name, e)
            raise

    def filter_exchange(self, stock_exchange_dict: dict[str, str]) -> str | None:
        """Choose ticker matching preferred exchange without '-' in symbol; else first clean key.

        Args:
            stock_exchange_dict (dict): Dictionary of stock exchanges (key: ticker, value: exchange)

        Returns:
            str: Filtered stock ticker based on preferred exchange
        """
        try:
            # Try to find preferred exchange first
            for key, value in stock_exchange_dict.items():
                if value == self.preferred_exchange and "-" not in key:
                    return key

            # If not found, return first valid key without '-'
            for key in stock_exchange_dict:
                if "-" not in key:
                    return key

            # If nothing matches, return None
            return None
        except Exception as e:
            self.log.error("Error in filter_exchange, %s", {e})
            raise

    # ------------------------- Data Fetch -------------------------

    def get_daily_returns_data(self, ticker, period, interval):
        """
        Fetches historical market data for a given stock ticker and computes daily returns.

        Args:
            ticker (str): Stock symbol/ticker (e.g. 'AAPL', 'GOOGL')
            period (str): Time period to download data for (e.g. '1d', '5d', '1mo', '1y')
            interval (str): Time interval between data points (e.g. '1m', '1h', '1d')

        Returns:
            pandas.DataFrame: DataFrame containing historical market data with percentage change in closing prices

        Note:
            Uses yfinance.download() to fetch market data from Yahoo Finance API
        """
        try:
            # Download historical market data
            data = yf.download(tickers=ticker, period=period, interval=interval)
            self.log.info(data.head())
            # Compute percentage change in closing prices
            data = data["Close"].pct_change()
            # Drop rows with NaN values
            data = data.dropna()
            data = data.iloc[::-1]
            return data
        except Exception as e:
            self.log.error(
                "Error in get_daily_returns_data, Ticker:%s, %s", {ticker}, {e}
            )
            raise

    # ------------------------- Metrics -------------------------

    def get_daily_returns_std_dev(self, ticker, period, interval):
        """Return standard deviation of daily returns as a plain float in a dict."""
        try:
            data = self.get_daily_returns_data(
                ticker=ticker, period=period, interval=interval
            )
            self.log.info(f"Daily returns data for {ticker}:\n{data.head()}")
            self.log.info(f"Data type: {type(data)}")
            data_std_dev = data.std()
            self.log.info(f"type of std dev: {type(data_std_dev)}")

            # If std() returned a Series (e.g., DataFrame.std()), extract the first value
            if isinstance(data_std_dev, pd.Series):
                value = round(float(data_std_dev.iloc[0]), 6)
            else:
                value = round(float(data_std_dev), 6)

            return {ticker: value}
        except Exception as e:
            self.log.error(
                "Error in get_daily_returns_std_dev, Ticker:%s, %s", {ticker}, {e}
            )
            raise

    def get_daily_returns_skewness(self, ticker, period, interval):
        """Return skewness of daily returns as a plain float in a dict."""
        try:
            data = self.get_daily_returns_data(
                ticker=ticker, period=period, interval=interval
            )
            print(f"Daily returns data for {ticker}:\n{data.head()}")
            print(f"Data type: {type(data)}")
            data_skewness = skew(data, bias=False)

            value = round(float(data_skewness), 6)

            return {ticker: value}
        except Exception as e:
            self.log.error(
                "Error in get_daily_returns_skewness, Ticker:%s, %s", {ticker}, {e}
            )
            raise

    def get_ticker_volume_mean(self, ticker, period, interval):
        """Return mean volume of the ticker as a plain float in a dict."""

        try:
            # Download historical market data
            data = yf.download(tickers=ticker, period=period, interval=interval)
            self.log.info(f"Volume data for {ticker}:\n{data['Volume'].head()}")
            self.log.info(f"Data type: {type(data)}")
            volume_mean = data["Volume"].mean()

            # Handle single-element Series safely to avoid FutureWarning
            if isinstance(volume_mean, pd.Series):
                scalar = float(volume_mean.iloc[0])
            else:
                scalar = float(volume_mean)

            value = round(scalar, 6)

            return {ticker: value}
        except Exception as e:
            self.log.error(
                "Error in get_ticker_volume_mean, Ticker:%s, %s", {ticker}, {e}
            )
            raise

    # ------------------------- Signals / Filters -------------------------

    def is_daily_returns_for_last_n_days(
        self, ticker, n_days="7d", interval="1d", positive_returns: bool = False
    ) -> bool:
        """
        Return True if the last `n_days` daily returns for `ticker` are all negative (or positive if specified).

        Notes:
        - Uses get_daily_returns_data() to fetch daily returns.
        """
        try:
            daily_returns = self.get_daily_returns_data(
                ticker=ticker, period=n_days, interval=interval
            )
            if daily_returns is None or daily_returns.empty:
                logging.info("No daily returns for %s", ticker)
                return False
            self.log.info(f"Daily returns for {ticker}:\n{daily_returns}")
            self.log.info(f"Data type: {type(daily_returns)}")

            # Normalize to a Series (handle single-column DataFrame or Series)
            if isinstance(daily_returns, pd.DataFrame):
                if daily_returns.shape[1] == 1:
                    series = daily_returns.iloc[:, 0]
                else:
                    # choose first column if multiple present (log for visibility)
                    logging.info(
                        "daily_returns has multiple columns; using first column '%s'",
                        daily_returns.columns[0],
                    )
                    series = daily_returns.iloc[:, 0]
            else:
                series = daily_returns

            self.log.info(f"Normalized series for {ticker}:\n{series}")
            self.log.info(f"Series type: {type(series)}")

            # Check all are positive if specified
            if positive_returns:
                return bool((series > 0).all())
            else:
                return bool((series < 0).all())

        except Exception as e:
            logging.error(
                "Error in is_daily_returns_for_last_n_days, Ticker:%s, %s",
                ticker,
                e,
            )
            raise

    # ------------------------- Excel I/O -------------------------
    def update_metrics_excel(
        self,
        excel_path: str,
        metric_name: str,
        metric_dict: dict[str, float],
        sheet_name: str = "Metrics",
        round_digits: int | None = 6,
    ):
        """
        Idempotently merge/update a {ticker: value} metric into an Excel sheet.
        - Preserves other sheets.
        - Creates file/sheet if missing.
        - Overwrites values for provided tickers.
        - Adds new tickers/columns as needed.
        """
        try:
            # 1) Build the incoming metric frame
            df_metric = pd.DataFrame.from_dict(
                metric_dict, orient="index", columns=[metric_name]
            )
            df_metric.index.name = "Ticker"
            if round_digits is not None:
                df_metric[metric_name] = df_metric[metric_name].round(round_digits)

            path = Path(excel_path)
            path.parent.mkdir(parents=True, exist_ok=True)

            # 2) Load existing sheet if present; else start with empty index named 'Ticker'
            if path.exists():
                try:
                    existing = pd.read_excel(excel_path, sheet_name=sheet_name)
                except ValueError:
                    existing = pd.DataFrame()
            else:
                existing = pd.DataFrame()

            # 3) Ensure we work with an index named 'Ticker' (no duplicate Ticker column)
            if not existing.empty:
                if "Ticker" not in existing.columns:
                    raise ValueError(
                        "Existing sheet is missing required 'Ticker' column."
                    )
                existing = existing.set_index("Ticker")
            else:
                # Start with an empty frame whose index is named 'Ticker'
                existing = pd.DataFrame(index=pd.Index([], name="Ticker"))

            # 4) Ensure all incoming tickers exist as rows (via union) and assign/overwrite values
            existing = existing.reindex(existing.index.union(df_metric.index))
            existing.loc[df_metric.index, metric_name] = df_metric[metric_name].values

            # 5) Write only this sheet, preserving other sheets in the workbook
            updated = existing.reset_index()
            if path.exists():
                with pd.ExcelWriter(
                    excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
                ) as writer:
                    updated.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
                    updated.to_excel(writer, sheet_name=sheet_name, index=False)

            return updated
        except Exception as e:
            logging.error(f"Error in update_metrics_excel, {e}")
            raise

    # ------------------------- Ranking -------------------------
    def rank_tickers_by_criteria_filtered(
        self,
        excel_path: str,
        sheet_name: str = "Metrics",
        ranking_sheet_name: str = "Ranking",
        criteria_weights: dict[str, float] = None,
        n_days: str = "7d",
        interval: str = "1d",
        ascending: bool = True,
        round_digits: int = 6,
        positive_returns: bool = False,
    ):
        """
        Rank tickers after filtering by skewness and consecutive negative daily returns.

        Filters:
            1. Include only tickers with skewness <= 0
            2. Include only tickers whose daily returns have been negative for the last `n_days`

        Ranking:
            - Each selected metric (e.g. Std Dev, Mean Volume) is converted to a rank
            - Combined rank = Weighted average of individual ranks
            - Lower combined rank = better

        Args:
            excel_path (str): Path to Excel containing metrics
            sheet_name (str): Source sheet name (default: "Metrics")
            ranking_sheet_name (str): Destination sheet for ranked results (default: "Ranking")
            criteria_weights (dict): e.g., {'Standard Deviation': 0.7, 'Mean Volume': 0.3}
            n_days (str): Period for negative-return filter (e.g. "7d")
            interval (str): Interval for return data (e.g. "1d")
            ascending (bool): True → smaller cumulative rank = better
            round_digits (int): Decimal precision for ranks
            positive_returns (bool): If True, filter for positive returns instead of negative
        """

        if not criteria_weights:
            raise ValueError(
                "criteria_weights must be provided, e.g. {'Standard Deviation': 0.7, 'Mean Volume': 0.3}"
            )

        # 1️⃣ Load Excel data
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if "Ticker" not in df.columns:
            raise ValueError("Excel must contain a 'Ticker' column.")
        df = df.set_index("Ticker")

        # 2️⃣ Ensure required columns exist
        required = list(criteria_weights.keys()) + ["Skewness"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns in Excel: {missing}")

        # 3️⃣ Filter step
        filtered_tickers = []
        for ticker, row in df.iterrows():
            skew_val = row.get("Skewness", 0)
            if pd.isna(skew_val) or skew_val > 0:
                continue  # skip if skew > 0

            try:
                if self.is_daily_returns_for_last_n_days(
                    ticker,
                    n_days=n_days,
                    interval=interval,
                    positive_returns=positive_returns,
                ):
                    filtered_tickers.append(ticker)
            except Exception as e:
                logging.warning(f"Skipping {ticker} due to error in return check: {e}")

        if not filtered_tickers:
            logging.warning("No tickers passed filters.")
            return pd.DataFrame()

        df_filtered = df.loc[filtered_tickers].copy()
        logging.info(
            f"{len(filtered_tickers)} tickers passed filters: {filtered_tickers}"
        )

        # 4️⃣ Assign ranks for each selected criterion
        for crit in criteria_weights.keys():
            # Determine ascending/descending based on metric meaning
            if "Volume" in crit:
                # higher volume = better → descending order
                df_filtered[f"Rank_{crit}"] = df_filtered[crit].rank(ascending=False)
            else:
                # higher std dev = better → descending order
                df_filtered[f"Rank_{crit}"] = df_filtered[crit].rank(ascending=False)

        # 5️⃣ Compute weighted composite rank
        total_weight = sum(criteria_weights.values())
        for crit in criteria_weights.keys():
            criteria_weights[crit] = criteria_weights[crit] / total_weight

        df_filtered["Cumulative_Rank"] = 0
        for crit, weight in criteria_weights.items():
            df_filtered["Cumulative_Rank"] += df_filtered[f"Rank_{crit}"] * weight

        df_filtered["Cumulative_Rank"] = df_filtered["Cumulative_Rank"].round(
            round_digits
        )

        # 6️⃣ Sort tickers based on final rank
        df_ranked = df_filtered.sort_values(
            "Cumulative_Rank", ascending=True
        ).reset_index()

        # 7️⃣ Select columns for output
        rank_cols = ["Ticker", "Cumulative_Rank"] + [
            c for c in df_filtered.columns if c.startswith("Rank_")
        ]
        metric_cols = list(criteria_weights.keys()) + ["Skewness"]
        df_ranked = df_ranked[["Ticker"] + metric_cols + rank_cols[1:]]

        # 8️⃣ Write ranked tickers to a new sheet in same workbook
        with pd.ExcelWriter(
            excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df_ranked.to_excel(writer, sheet_name=ranking_sheet_name, index=False)

        return df_ranked

    # ------------------------- Batch Orchestration -------------------------
    def process_constituents(
        self,
        input_excel_path: str,
        output_excel_path: str,
        period: str = "5d",
        interval: str = "1d",
        n_days: str = "7d",
        criteria_weights: dict | None = None,
        positive_returns: bool = False,
    ):
        """
        Loop over names in input_excel_path['Name'], resolve tickers, compute metrics and
        write them to output_excel_path (Metrics sheet). Optionally produce ranked sheet.
        """
        try:
            # read input list
            df_in = pd.read_excel(input_excel_path)
            if "Name" not in df_in.columns:
                raise ValueError("Input excel must contain a 'Name' column.")
            names = df_in["Name"].dropna().astype(str).unique()

            # ensure output directory exists
            Path(output_excel_path).parent.mkdir(parents=True, exist_ok=True)

            for name in names:
                try:
                    logging.info("Processing company: %s", name)
                    stock_exchanges = self.lookup_stock(name=name)
                    if not stock_exchanges:
                        logging.warning("No lookup result for '%s'", name)
                        continue

                    ticker = self.filter_exchange(stock_exchanges)
                    if not ticker:
                        logging.warning(
                            "No suitable ticker found for '%s' (lookup=%s)",
                            name,
                            stock_exchanges,
                        )
                        continue

                    # compute metrics (each returns {ticker: value})
                    std_dev = self.get_daily_returns_std_dev(
                        ticker=ticker, period=period, interval=interval
                    )
                    vol_mean = self.get_ticker_volume_mean(
                        ticker=ticker, period=period, interval=interval
                    )
                    skewness = self.get_daily_returns_skewness(
                        ticker=ticker, period=period, interval=interval
                    )

                    # update Excel per metric (idempotent)
                    self.update_metrics_excel(
                        excel_path=output_excel_path,
                        metric_name="Standard Deviation",
                        metric_dict=std_dev,
                    )
                    self.update_metrics_excel(
                        excel_path=output_excel_path,
                        metric_name="Volume",
                        metric_dict=vol_mean,
                    )
                    self.update_metrics_excel(
                        excel_path=output_excel_path,
                        metric_name="Skewness",
                        metric_dict=skewness,
                    )

                    logging.info("Finished %s -> %s", name, ticker)

                except Exception as item_err:
                    logging.error("Failed processing '%s': %s", name, item_err)

            # produce ranked sheet if weights provided
            if criteria_weights:
                logging.info(
                    "Generating ranking sheet with weights: %s", criteria_weights
                )
                ranked = self.rank_tickers_by_criteria_filtered(
                    excel_path=output_excel_path,
                    sheet_name="Metrics",
                    ranking_sheet_name="Filtered_Ranking",
                    criteria_weights=criteria_weights,
                    n_days=n_days,
                    interval=interval,
                    ascending=False,
                    positive_returns=positive_returns,
                )
                return ranked

            return None

        except Exception as e:
            logging.error("Error in process_constituents: %s", e)
            raise


if __name__ == "__main__":
    # 1️⃣ Ensure the directory exists (create if not)
    os.makedirs("outputs", exist_ok=True)
    # input output file paths
    input_file = os.path.join("input", "WLS_Constituents.xlsx")
    output_file = os.path.join("outputs", "screen_metrics.xlsx")
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")

    # parameter file (edit this file to change run parameters)
    params_path = Path("parameter.json")

    # load parameters
    with params_path.open("r", encoding="utf-8") as f:
        params = json.load(f)

    period = params.get("period", "5d")
    interval = params.get("interval", "1d")
    n_days = params.get("n_days", "7d")
    criteria_weights = params.get("criteria_weights")
    preferred_exchange = params.get("preferred_exchange", EXCHANGE)
    filter_positive_returns = params.get("filter_positive_returns", False)

    print(f"Period: {period}, Interval: {interval}, n_days: {n_days}")
    print(f"Criteria weights: {criteria_weights}")
    print(f"Preferred exchange: {preferred_exchange}")

    screener = StockScreener(preferred_exchange=preferred_exchange)

    # run
    # ranked_df = screener.process_constituents(
    #     input_excel_path=input_file,
    #     output_excel_path=output_file,
    #     period="5d",
    #     interval="1d",
    #     n_days="7d",
    #     criteria_weights=criteria_weights,
    # )
    # if ranked_df is not None:
    #     print("Ranking completed. Top rows:")
    #     print(ranked_df.head())

    # excel_file_path = os.path.join("outputs", "screen_metrics.xlsx")

    # screener = StockScreener(preferred_exchange=EXCHANGE)

    # Ranking directly from existing Metrics sheet
    ranked_df = screener.rank_tickers_by_criteria_filtered(
        excel_path=output_file,
        sheet_name="Metrics",
        ranking_sheet_name="Filtered_Ranking",
        criteria_weights=criteria_weights,
        n_days=n_days,
        interval=interval,
        ascending=False,  # keep current behavior
        positive_returns=filter_positive_returns,
    )
    print(ranked_df)
