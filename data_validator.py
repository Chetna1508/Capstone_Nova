
from __future__ import annotations

import pandas as pd


EXPECTED_COLUMNS = {
    "Regional_Performance": {
        "Quarter",
        "Region",
        "Revenue_USD_M",
        "Units_Sold_K",
        "Returns_Rate_pct",
        "Marketing_Spend_USD_M",
    },
    "Customer_Experience": {
        "Quarter",
        "Region",
        "NPS",
        "CSAT_pct",
        "Avg_Delivery_Days",
        "Support_Tickets_K",
    },
    "Operational_Risk": {
        "Quarter",
        "Region",
        "Stockout_Rate_pct",
        "Return_Processing_Days",
        "Warehouse_Capacity_Util_pct",
        "Fraud_Incidents",
    },
}


def validate_data(
    dataframes: dict[str, pd.DataFrame],
) -> None:
    """Raise a clear error when required sheets or values are invalid."""

    missing_sheets = set(EXPECTED_COLUMNS) - set(dataframes)

    if missing_sheets:
        raise ValueError(
            "Missing required worksheets: "
            f"{sorted(missing_sheets)}"
        )

    keys_by_sheet: dict[
        str,
        set[tuple[object, object]],
    ] = {}

    for sheet_name, expected_columns in EXPECTED_COLUMNS.items():
        dataframe = dataframes[sheet_name]

        missing_columns = expected_columns - set(
            dataframe.columns
        )

        if missing_columns:
            raise ValueError(
                f"{sheet_name} is missing columns: "
                f"{sorted(missing_columns)}"
            )

        if dataframe.empty:
            raise ValueError(
                f"{sheet_name} contains no data."
            )

        if (
            dataframe[list(expected_columns)]
            .isna()
            .any()
            .any()
        ):
            raise ValueError(
                f"{sheet_name} contains missing values "
                "in required columns."
            )

        duplicate_count = int(
            dataframe.duplicated(
                subset=["Quarter", "Region"]
            ).sum()
        )

        if duplicate_count:
            raise ValueError(
                f"{sheet_name} has {duplicate_count} "
                "duplicate Quarter/Region rows."
            )

        keys_by_sheet[sheet_name] = set(
            dataframe[
                ["Quarter", "Region"]
            ].itertuples(
                index=False,
                name=None,
            )
        )

    baseline_name = "Regional_Performance"
    baseline_keys = keys_by_sheet[baseline_name]

    for sheet_name, keys in keys_by_sheet.items():
        if keys != baseline_keys:
            missing = sorted(
                baseline_keys - keys
            )
            extra = sorted(
                keys - baseline_keys
            )

            raise ValueError(
                f"{sheet_name} does not align with "
                f"{baseline_name}. Missing keys: "
                f"{missing[:5]}; extra keys: {extra[:5]}"
            )

    quarter_count = dataframes[
        baseline_name
    ]["Quarter"].nunique()

    if quarter_count < 2:
        raise ValueError(
            "At least two quarters are required "
            "for QoQ analysis."
        )