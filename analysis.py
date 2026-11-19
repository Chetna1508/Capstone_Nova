
from __future__ import annotations

import re
from typing import Any

import pandas as pd


def _quarter_sort_key(
    value: object,
) -> tuple[int, int, str]:
    text = str(value).strip()

    match = re.search(
        r"Q([1-4])\D*(\d{4})",
        text,
        re.IGNORECASE,
    )

    if match:
        return (
            int(match.group(2)),
            int(match.group(1)),
            text,
        )

    match = re.search(
        r"(\d{4})\D*Q([1-4])",
        text,
        re.IGNORECASE,
    )

    if match:
        return (
            int(match.group(1)),
            int(match.group(2)),
            text,
        )

    match = re.fullmatch(
        r"Q([1-4])",
        text,
        re.IGNORECASE,
    )

    if match:
        return (
            0,
            int(match.group(1)),
            text,
        )

    return 0, 0, text


def get_sorted_quarters(
    dataframe: pd.DataFrame,
) -> list[object]:
    """Return unique quarter labels in chronological order."""

    return sorted(
        dataframe[
            "Quarter"
        ].dropna().unique().tolist(),
        key=_quarter_sort_key,
    )


def get_latest_quarter(
    dataframes: dict[str, pd.DataFrame],
) -> object:
    return get_sorted_quarters(
        dataframes["Regional_Performance"]
    )[-1]


def _weighted_nps(
    performance: pd.DataFrame,
    customer: pd.DataFrame,
    quarter: object,
) -> float:
    performance_q = performance.loc[
        performance["Quarter"] == quarter,
        ["Region", "Units_Sold_K"],
    ]

    customer_q = customer.loc[
        customer["Quarter"] == quarter,
        ["Region", "NPS"],
    ]

    merged = customer_q.merge(
        performance_q,
        on="Region",
        how="inner",
        validate="one_to_one",
    )

    weight = float(
        merged["Units_Sold_K"].sum()
    )

    if weight <= 0:
        raise ValueError(
            "Units sold must be positive to calculate "
            f"weighted NPS for {quarter}."
        )

    return float(
        (
            merged["NPS"]
            * merged["Units_Sold_K"]
        ).sum()
        / weight
    )


def calculate_kpis(
    dataframes: dict[str, pd.DataFrame],
) -> dict[str, Any]:
    """
    Calculate verified headline KPIs.

    The ADK agent interprets these values but does not
    recalculate them.
    """

    performance = dataframes[
        "Regional_Performance"
    ]

    customer = dataframes[
        "Customer_Experience"
    ]

    quarters = get_sorted_quarters(
        performance
    )

    if len(quarters) < 2:
        raise ValueError(
            "At least two quarters are required "
            "for KPI analysis."
        )

    previous_quarter, latest_quarter = (
        quarters[-2:]
    )

    latest_revenue = float(
        performance.loc[
            performance["Quarter"]
            == latest_quarter,
            "Revenue_USD_M",
        ].sum()
    )

    previous_revenue = float(
        performance.loc[
            performance["Quarter"]
            == previous_quarter,
            "Revenue_USD_M",
        ].sum()
    )

    if previous_revenue == 0:
        raise ValueError(
            "Previous-quarter revenue cannot be zero."
        )

    return {
        "latest_quarter": str(
            latest_quarter
        ),
        "previous_quarter": str(
            previous_quarter
        ),
        "total_revenue_usd_m": (
            latest_revenue
        ),
        "qoq_revenue_growth_pct": (
            (
                latest_revenue
                - previous_revenue
            )
            / previous_revenue
            * 100
        ),
        "weighted_nps": _weighted_nps(
            performance,
            customer,
            latest_quarter,
        ),
    }


def analyze_regional_performance(
    performance: pd.DataFrame,
) -> dict[str, Any]:
    """
    Identify regional growth and returns highlights
    deterministically.
    """

    quarters = get_sorted_quarters(
        performance
    )

    first_quarter = quarters[0]
    latest_quarter = quarters[-1]

    first = performance.loc[
        performance["Quarter"]
        == first_quarter,
        ["Region", "Revenue_USD_M"],
    ].rename(
        columns={
            "Revenue_USD_M": "First_Revenue"
        }
    )

    latest = performance.loc[
        performance["Quarter"]
        == latest_quarter
    ].copy()

    comparison = latest.merge(
        first,
        on="Region",
        validate="one_to_one",
    )

    if (
        comparison["First_Revenue"]
        == 0
    ).any():
        raise ValueError(
            "First-quarter regional revenue "
            "cannot be zero."
        )

    comparison[
        "Revenue_Change_pct"
    ] = (
        (
            comparison["Revenue_USD_M"]
            - comparison["First_Revenue"]
        )
        / comparison["First_Revenue"]
        * 100
    )

    strongest = comparison.loc[
        comparison[
            "Revenue_Change_pct"
        ].idxmax()
    ]

    weakest = comparison.loc[
        comparison[
            "Revenue_Change_pct"
        ].idxmin()
    ]

    highest_returns = comparison.loc[
        comparison[
            "Returns_Rate_pct"
        ].idxmax()
    ]

    return {
        "comparison": comparison,
        "strongest_growth_region": str(
            strongest["Region"]
        ),
        "strongest_growth_pct": float(
            strongest[
                "Revenue_Change_pct"
            ]
        ),
        "weakest_growth_region": str(
            weakest["Region"]
        ),
        "weakest_growth_pct": float(
            weakest[
                "Revenue_Change_pct"
            ]
        ),
        "highest_returns_region": str(
            highest_returns["Region"]
        ),
        "highest_returns_pct": float(
            highest_returns[
                "Returns_Rate_pct"
            ]
        ),
    }


def _records(
    dataframe: pd.DataFrame,
    columns: list[str],
) -> list[dict[str, object]]:
    records: list[
        dict[str, object]
    ] = []

    raw_records = (
        dataframe[columns]
        .sort_values("Region")
        .to_dict("records")
    )

    for raw in raw_records:
        record: dict[
            str,
            object,
        ] = {}

        for key, value in raw.items():
            if hasattr(value, "item"):
                value = value.item()

            record[key] = value

        records.append(record)

    return records


def build_business_snapshot(
    dataframes: dict[str, pd.DataFrame],
) -> dict[str, Any]:
    """
    Build the JSON-safe fact package returned by
    the ADK function tool.
    """

    performance = dataframes[
        "Regional_Performance"
    ]

    customer = dataframes[
        "Customer_Experience"
    ]

    risk = dataframes[
        "Operational_Risk"
    ]

    kpis = calculate_kpis(
        dataframes
    )

    regional = (
        analyze_regional_performance(
            performance
        )
    )

    latest_label = kpis[
        "latest_quarter"
    ]

    latest_performance = performance[
        performance[
            "Quarter"
        ].astype(str)
        == latest_label
    ]

    latest_customer = customer[
        customer[
            "Quarter"
        ].astype(str)
        == latest_label
    ]

    latest_risk = risk[
        risk[
            "Quarter"
        ].astype(str)
        == latest_label
    ]

    quarterly_trend = []

    for quarter in get_sorted_quarters(
        performance
    ):
        revenue = float(
            performance.loc[
                performance["Quarter"]
                == quarter,
                "Revenue_USD_M",
            ].sum()
        )

        quarterly_trend.append(
            {
                "quarter": str(
                    quarter
                ),
                "revenue_usd_m": round(
                    revenue,
                    2,
                ),
                "sales_weighted_nps": round(
                    _weighted_nps(
                        performance,
                        customer,
                        quarter,
                    ),
                    2,
                ),
            }
        )

    return {
        "project": (
            "NovaRetail Quarterly "
            "Business Review"
        ),
        "period": latest_label,
        "headline_kpis": {
            "total_revenue_usd_m": round(
                float(
                    kpis[
                        "total_revenue_usd_m"
                    ]
                ),
                2,
            ),
            "qoq_revenue_growth_pct": round(
                float(
                    kpis[
                        "qoq_revenue_growth_pct"
                    ]
                ),
                2,
            ),
            "sales_weighted_nps": round(
                float(
                    kpis[
                        "weighted_nps"
                    ]
                ),
                2,
            ),
        },
        "quarterly_trend": (
            quarterly_trend
        ),
        "regional_highlights": {
            "strongest_growth_region": (
                regional[
                    "strongest_growth_region"
                ]
            ),
            "strongest_growth_pct": round(
                float(
                    regional[
                        "strongest_growth_pct"
                    ]
                ),
                2,
            ),
            "weakest_growth_region": (
                regional[
                    "weakest_growth_region"
                ]
            ),
            "weakest_growth_pct": round(
                float(
                    regional[
                        "weakest_growth_pct"
                    ]
                ),
                2,
            ),
            "highest_returns_region": (
                regional[
                    "highest_returns_region"
                ]
            ),
            "highest_returns_pct": round(
                float(
                    regional[
                        "highest_returns_pct"
                    ]
                ),
                2,
            ),
        },
        "latest_regional_performance": (
            _records(
                latest_performance,
                [
                    "Region",
                    "Revenue_USD_M",
                    "Units_Sold_K",
                    "Returns_Rate_pct",
                    "Marketing_Spend_USD_M",
                ],
            )
        ),
        "latest_customer_experience": (
            _records(
                latest_customer,
                [
                    "Region",
                    "NPS",
                    "CSAT_pct",
                    "Avg_Delivery_Days",
                    "Support_Tickets_K",
                ],
            )
        ),
        "latest_operational_risk": (
            _records(
                latest_risk,
                [
                    "Region",
                    "Stockout_Rate_pct",
                    "Return_Processing_Days",
                    "Warehouse_Capacity_Util_pct",
                    "Fraud_Incidents",
                ],
            )
        ),
    }