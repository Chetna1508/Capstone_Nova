from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation

from .agent import generate_executive_insights
from .src.analysis import (
    build_business_snapshot,
    calculate_kpis,
)
from .src.chart_generation import (
    create_customer_chart,
    create_performance_chart,
    create_risk_chart,
)
from .src.config import (
    ALLOW_AGENT_FALLBACK,
    DATASET_PATH,
    ICON_MAPPING,
    OUTPUT_DIR,
    OUTPUT_PPTX,
    TEMPLATE_PATH,
)
from .src.content_check import (
    check_instruction_text,
)
from .src.data_loader import (
    load_data,
)
from .src.data_validator import (
    validate_data,
)
from .src.ppt_builder import (
    remove_template_instructions,
    replace_icon_if_available,
    replace_placeholder_with_image,
    set_bullets,
    set_text,
)


def _format_revenue(
    value_usd_m: float,
) -> str:
    if abs(value_usd_m) >= 1000:
        return (
            f"${value_usd_m / 1000:.2f}B"
        )

    return (
        f"${value_usd_m:,.0f}M"
    )


def _set_optional_text(
    slide,
    names: tuple[str, ...],
    text: str,
) -> None:
    try:
        set_text(
            slide,
            names,
            text,
        )
    except KeyError:
        # Some templates already contain fixed titles.
        return


def _add_icons(
    presentation: Presentation,
) -> None:
    for index, slide in enumerate(
        presentation.slides,
        start=1,
    ):
        icon_path = ICON_MAPPING.get(
            index
        )

        if icon_path is None:
            continue

        replace_icon_if_available(
            slide,
            (
                f"icon_{index}",
                f"icon{index}",
                f"Icon {index}",
            ),
            icon_path,
        )


def build_presentation(
    *,
    template_path: str | Path,
    output_path: str | Path,
    kpis: dict[str, Any],
    insights: dict[str, Any],
    chart_paths: dict[str, Path],
) -> Path:
    """
    Populate the supplied five-slide template with deterministic
    facts and ADK-generated narrative.
    """

    template = Path(
        template_path
    )

    if not template.is_file():
        raise FileNotFoundError(
            "NovaRetail template not found: "
            f"{template.resolve()}"
        )

    presentation = Presentation(
        str(template)
    )

    if len(
        presentation.slides
    ) != 5:
        raise ValueError(
            "The NovaRetail template must contain exactly "
            "five slides; found "
            f"{len(presentation.slides)}."
        )

    (
        slide_1,
        slide_2,
        slide_3,
        slide_4,
        slide_5,
    ) = presentation.slides

    # Slide 1: headline metrics and executive story.
    set_text(
        slide_1,
        (
            "subtitle_1",
            "subtitle1",
            "Subtitle 1",
        ),
        (
            f"{kpis['latest_quarter']} "
            "Quarterly Business Review"
        ),
    )

    set_text(
        slide_1,
        (
            "kpi1_value",
            "kpi_1_value",
        ),
        _format_revenue(
            float(
                kpis[
                    "total_revenue_usd_m"
                ]
            )
        ),
    )

    set_text(
        slide_1,
        (
            "kpi1_label",
            "kpi_1_label",
        ),
        "Revenue",
    )

    set_text(
        slide_1,
        (
            "kpi2_value",
            "kpi_2_value",
        ),
        (
            f"{float(kpis['qoq_revenue_growth_pct']):+.1f}%"
        ),
    )

    set_text(
        slide_1,
        (
            "kpi2_label",
            "kpi_2_label",
        ),
        "QoQ Revenue Growth",
    )

    set_text(
        slide_1,
        (
            "kpi3_value",
            "kpi_3_value",
        ),
        (
            f"{float(kpis['weighted_nps']):.1f}"
        ),
    )

    set_text(
        slide_1,
        (
            "kpi3_label",
            "kpi_3_label",
        ),
        "Sales-weighted NPS",
    )

    set_text(
        slide_1,
        (
            "summary_text",
            "executive_summary",
            "executive_summary_text",
        ),
        insights[
            "executive_summary"
        ],
    )

    # Slide 2: revenue and regional performance.
    _set_optional_text(
        slide_2,
        (
            "title_2",
            "performance_title",
            "title_performance",
        ),
        "Performance & Regional Momentum",
    )

    replace_placeholder_with_image(
        slide_2,
        (
            "chart_performance",
            "chart_placeholder_2",
            "performance_chart_placeholder",
            "performance_chart",
            "chart_2",
            "Chart 2",
        ),
        chart_paths[
            "performance"
        ],
    )

    set_bullets(
        slide_2,
        (
            "insights_performance",
            "performance_text",
            "performance_insights",
            "insights_text_2",
            "body_2",
        ),
        insights[
            "performance_bullets"
        ],
    )

    # Slide 3: customer experience.
    _set_optional_text(
        slide_3,
        (
            "title_3",
            "customer_title",
            "title_customer",
        ),
        "Customer Experience",
    )

    replace_placeholder_with_image(
        slide_3,
        (
            "chart_customer",
            "chart_placeholder_3",
            "customer_chart_placeholder",
            "customer_chart",
            "chart_3",
            "Chart 3",
        ),
        chart_paths[
            "customer"
        ],
    )

    set_bullets(
        slide_3,
        (
            "insights_customer",
            "customer_text",
            "customer_insights",
            "insights_text_3",
            "body_3",
        ),
        insights[
            "customer_bullets"
        ],
    )

    # Slide 4: operational risk.
    _set_optional_text(
        slide_4,
        (
            "title_4",
            "risk_title",
            "title_risk",
        ),
        "Operational Risk",
    )

    replace_placeholder_with_image(
        slide_4,
        (
            "chart_risk",
            "chart_placeholder_4",
            "risk_chart_placeholder",
            "risk_chart",
            "chart_4",
            "Chart 4",
        ),
        chart_paths[
            "risk"
        ],
    )

    set_bullets(
        slide_4,
        (
            "insights_risk",
            "risk_text",
            "risk_insights",
            "insights_text_4",
            "body_4",
        ),
        insights[
            "risk_bullets"
        ],
    )

    # Slide 5: evidence-linked recommendations.
    _set_optional_text(
        slide_5,
        (
            "title_5",
            "recommendations_title",
            "actions_title",
            "title_recommendations",
        ),
        "Recommended Actions",
    )

    set_bullets(
        slide_5,
        (
            "recommendations_text",
            "insights_recommendations",
            "recommendation_text",
            "actions_text",
            "body_5",
        ),
        insights[
            "recommendations"
        ],
    )

    set_text(
        slide_5,
        (
            "closing_statement",
            "closing_text",
            "takeaway_text",
        ),
        insights[
            "closing_statement"
        ],
    )

    _add_icons(
        presentation
    )

    for slide in presentation.slides:
        remove_template_instructions(
            slide
        )

    output = Path(
        output_path
    )

    output.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    presentation.save(
        str(output)
    )

    check_instruction_text(
        output
    )

    return output


def run_pipeline() -> Path:
    """Execute the complete automated QBR workflow once."""

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    dataframes = load_data(
        DATASET_PATH
    )

    validate_data(
        dataframes
    )

    kpis = calculate_kpis(
        dataframes
    )

    snapshot = build_business_snapshot(
        dataframes
    )

    # Required intelligent component: the agent consumes
    # deterministic KPI, weighted-NPS, customer, and risk outputs.
    insights = generate_executive_insights(
        snapshot,
        allow_fallback=(
            ALLOW_AGENT_FALLBACK
        ),
    )

    chart_paths = {
        "performance": (
            create_performance_chart(
                dataframes[
                    "Regional_Performance"
                ]
            )
        ),
        "customer": (
            create_customer_chart(
                dataframes[
                    "Customer_Experience"
                ]
            )
        ),
        "risk": (
            create_risk_chart(
                dataframes[
                    "Operational_Risk"
                ]
            )
        ),
    }

    return build_presentation(
        template_path=TEMPLATE_PATH,
        output_path=OUTPUT_PPTX,
        kpis=kpis,
        insights=insights,
        chart_paths=chart_paths,
    )


def main() -> None:
    output = run_pipeline()

    print(
        "NovaRetail QBR created successfully: "
        f"{output.resolve()}"
    )


if __name__ == "__main__":
    main()