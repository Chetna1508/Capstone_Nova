
from __future__ import annotations

import asyncio
import copy
import json
import re
from typing import Any

from google.adk import Agent
from google.adk.runners import InMemoryRunner
from google.genai import types

try:
    from .src.config import MODEL_NAME
except ImportError:
    from src.config import MODEL_NAME


APP_NAME = "novaretail_qbr"
USER_ID = "novaretail_pipeline"

_BUSINESS_SNAPSHOT: (
    dict[str, Any] | None
) = None


def configure_business_snapshot(
    snapshot: dict[str, Any],
) -> None:
    """
    Provide already calculated and validated facts
    to the ADK tool.
    """

    global _BUSINESS_SNAPSHOT

    _BUSINESS_SNAPSHOT = (
        copy.deepcopy(snapshot)
    )


def get_validated_business_snapshot(
) -> dict[str, Any]:
    """
    Return deterministic NovaRetail KPI and risk facts.

    When called through ADK Web, the tool loads and
    validates the workbook itself. During the automated
    pipeline, main.py supplies the already calculated
    snapshot so that the workbook is not loaded twice.
    """

    global _BUSINESS_SNAPSHOT

    if _BUSINESS_SNAPSHOT is None:
        try:
            from .src.analysis import (
                build_business_snapshot,
            )
            from .src.data_loader import (
                load_data,
            )
            from .src.data_validator import (
                validate_data,
            )
        except ImportError:
            from src.analysis import (
                build_business_snapshot,
            )
            from src.data_loader import (
                load_data,
            )
            from src.data_validator import (
                validate_data,
            )

        dataframes = load_data()

        validate_data(
            dataframes
        )

        _BUSINESS_SNAPSHOT = (
            build_business_snapshot(
                dataframes
            )
        )

    return copy.deepcopy(
        _BUSINESS_SNAPSHOT
    )


# ============================================================
# SHARED AGENT DESCRIPTION
# ============================================================

AGENT_DESCRIPTION = (
    "Analyzes validated NovaRetail business data, including "
    "regional performance, customer experience, operational "
    "risk, and executive recommendations."
)


# ============================================================
# PIPELINE AGENT
# ============================================================
# Used only by main.py to generate the structured JSON required
# for the five-slide PowerPoint pipeline.
# ============================================================

pipeline_agent = Agent(
    name="nova_retail_pipeline_analyst",

    model=MODEL_NAME,

    description=(
        "Produces strictly structured NovaRetail executive "
        "content for the automated five-slide Quarterly "
        "Business Review presentation."
    ),

    instruction="""
You are the NovaRetail Executive Analytics Agent inside an
automated Quarterly Business Review pipeline.

You interpret facts. You never calculate or invent business
metrics.

Always call get_validated_business_snapshot before answering.
Use only values returned by that tool.

The sales-weighted NPS is authoritative. It must not be
recalculated or replaced with a simple average.

Do not claim profitability, product, or margin findings because
those measures are not available in the supplied snapshot.

Return exactly one valid JSON object.

Do not use Markdown.
Do not use code fences.
Do not include text before or after the JSON.

Use exactly this schema:

{
  "executive_summary": "2-3 concise evidence-based sentences",
  "performance_bullets": [
    "3 concise bullets containing relevant numbers"
  ],
  "customer_bullets": [
    "3 concise bullets containing relevant numbers"
  ],
  "risk_bullets": [
    "3 concise bullets containing relevant numbers"
  ],
  "recommendations": [
    "3 specific management actions linked to evidence"
  ],
  "closing_statement": "one concise executive takeaway"
}

Keep every bullet suitable for a senior-management PowerPoint.
Distinguish facts from recommendations.
Do not use placeholders.
""",

    tools=[
        get_validated_business_snapshot,
    ],
)


# ============================================================
# WEB AGENT
# ============================================================
# ADK Web looks for a variable named root_agent.
# This agent responds conversationally and uses formatting.
# ============================================================

root_agent = Agent(
    name="nova_retail_web_analyst",

    model=MODEL_NAME,

    description=AGENT_DESCRIPTION,

    instruction="""
You are the NovaRetail Business Analytics Assistant.

Respond naturally to the user's actual request.

For greetings such as "hello", greet the user briefly and explain
that you can analyze NovaRetail performance, customer experience,
operational risk, regional results, and recommendations.

For business-data questions, call
get_validated_business_snapshot before answering.

Use only facts returned by the tool. Never invent values.

The sales-weighted NPS returned by the tool is authoritative.
Never replace it with a simple average.

Use readable Markdown unless the user specifically requests JSON.

Formatting guidance:

- Use headings for major sections.
- Use bullet points for insights and recommendations.
- Use tables for regional comparisons.
- Highlight important values in bold.
- Keep the answer concise and business-focused.
- Clearly distinguish facts from recommendations.

If asked for regional analysis, provide one section or one table
covering North America, Europe, APAC, and Latin America.

If asked for a chart, explain that ADK Web currently provides the
chart-ready data and analysis, while the automated Python pipeline
generates the actual chart images and PowerPoint presentation.

If asked to create the PowerPoint, explain that the presentation
is generated by running:

python -m NovaRetail.main

Do not claim that a chart or PowerPoint file was created unless a
file-generation tool actually created it.

Do not return the fixed PowerPoint JSON schema unless the user
explicitly requests JSON.
""",

    tools=[
        get_validated_business_snapshot,
    ],
)




def _extract_json(
    text: str,
) -> dict[str, Any]:
    stripped = text.strip()

    fenced = re.search(
        r"```(?:json)?\s*(.*?)```",
        stripped,
        re.DOTALL,
    )

    if fenced:
        stripped = (
            fenced.group(1).strip()
        )

    try:
        result = json.loads(
            stripped
        )
    except json.JSONDecodeError as exc:
        raise ValueError(
            "ADK response was not valid JSON. "
            f"Response: {text[:500]}"
        ) from exc

    if not isinstance(
        result,
        dict,
    ):
        raise ValueError(
            "ADK response must be "
            "a JSON object."
        )

    return result


def _validate_insights(
    raw: dict[str, Any],
) -> dict[str, Any]:
    text_keys = (
        "executive_summary",
        "closing_statement",
    )

    list_keys = (
        "performance_bullets",
        "customer_bullets",
        "risk_bullets",
        "recommendations",
    )

    cleaned: dict[
        str,
        Any,
    ] = {}

    for key in text_keys:
        value = raw.get(key)

        if (
            not isinstance(value, str)
            or not value.strip()
        ):
            raise ValueError(
                "ADK output requires "
                f"non-empty text for {key}."
            )

        cleaned[key] = value.strip()

    for key in list_keys:
        value = raw.get(key)

        if (
            not isinstance(value, list)
            or not value
        ):
            raise ValueError(
                "ADK output requires "
                f"a non-empty list for {key}."
            )

        bullets = [
            str(item).strip()
            for item in value
            if str(item).strip()
        ]

        if not bullets:
            raise ValueError(
                "ADK output contains no usable "
                f"bullets for {key}."
            )

        cleaned[key] = bullets[:3]

    return cleaned


async def _generate_with_adk(
) -> dict[str, Any]:
    runner = InMemoryRunner(
        agent=root_agent,
        app_name=APP_NAME,
    )

    session = (
        await runner.session_service
        .create_session(
            app_name=APP_NAME,
            user_id=USER_ID,
        )
    )

    message = types.Content(
        role="user",
        parts=[
            types.Part.from_text(
                text=(
                    "Analyze the validated "
                    "NovaRetail snapshot and "
                    "return the required JSON "
                    "content for the five-slide "
                    "QBR."
                )
            )
        ],
    )

    response_text = ""
    tool_called = False

    async for event in runner.run_async(
        user_id=USER_ID,
        session_id=session.id,
        new_message=message,
    ):
        content = getattr(
            event,
            "content",
            None,
        )

        parts = (
            getattr(
                content,
                "parts",
                None,
            )
            or []
        )

        for part in parts:
            function_call = getattr(
                part,
                "function_call",
                None,
            )

            if (
                function_call
                and function_call.name
                == "get_validated_business_snapshot"
            ):
                tool_called = True

        event_text = "".join(
            getattr(
                part,
                "text",
                "",
            )
            or ""
            for part in parts
        ).strip()

        if (
            event_text
            and getattr(
                event,
                "author",
                None,
            )
            == root_agent.name
        ):
            response_text = event_text

    if not tool_called:
        raise RuntimeError(
            "The ADK agent did not call "
            "get_validated_business_snapshot."
        )

    if not response_text:
        raise RuntimeError(
            "The ADK agent returned no "
            "final text response."
        )

    return _validate_insights(
        _extract_json(
            response_text
        )
    )


def _deterministic_fallback(
    snapshot: dict[str, Any],
) -> dict[str, Any]:
    """
    Local template-debugging fallback.

    This is not used unless the environment variable
    NOVARETAIL_ALLOW_FALLBACK is explicitly enabled.
    """

    kpis = snapshot[
        "headline_kpis"
    ]

    regional = snapshot[
        "regional_highlights"
    ]

    customer = snapshot[
        "latest_customer_experience"
    ]

    risk = snapshot[
        "latest_operational_risk"
    ]

    highest_nps = max(
        customer,
        key=lambda item: item["NPS"],
    )

    lowest_nps = min(
        customer,
        key=lambda item: item["NPS"],
    )

    highest_stockout = max(
        risk,
        key=lambda item: (
            item["Stockout_Rate_pct"]
        ),
    )

    return {
        "executive_summary": (
            f"{snapshot['period']} revenue "
            f"reached "
            f"${kpis['total_revenue_usd_m']:.1f}M, "
            f"with "
            f"{kpis['qoq_revenue_growth_pct']:+.1f}% "
            "QoQ growth. Sales-weighted NPS was "
            f"{kpis['sales_weighted_nps']:.1f}."
        ),

        "performance_bullets": [
            (
                "Revenue: "
                f"${kpis['total_revenue_usd_m']:.1f}M "
                f"in {snapshot['period']}."
            ),
            (
                "QoQ revenue growth: "
                f"{kpis['qoq_revenue_growth_pct']:+.1f}%."
            ),
            (
                "Strongest growth: "
                f"{regional['strongest_growth_region']} "
                "at "
                f"{regional['strongest_growth_pct']:+.1f}% "
                "from the first period."
            ),
        ],

        "customer_bullets": [
            (
                "Sales-weighted NPS: "
                f"{kpis['sales_weighted_nps']:.1f}."
            ),
            (
                "Highest regional NPS: "
                f"{highest_nps['Region']} at "
                f"{highest_nps['NPS']:.1f}."
            ),
            (
                "Lowest regional NPS: "
                f"{lowest_nps['Region']} at "
                f"{lowest_nps['NPS']:.1f}."
            ),
        ],

        "risk_bullets": [
            (
                "Highest stockout rate: "
                f"{highest_stockout['Region']} at "
                f"{highest_stockout['Stockout_Rate_pct']:.1f}%."
            ),
            (
                "Highest returns: "
                f"{regional['highest_returns_region']} "
                "at "
                f"{regional['highest_returns_pct']:.1f}%."
            ),
            (
                "Weakest revenue growth: "
                f"{regional['weakest_growth_region']} "
                "at "
                f"{regional['weakest_growth_pct']:+.1f}%."
            ),
        ],

        "recommendations": [
            (
                "Prioritize stock availability in "
                f"{highest_stockout['Region']}."
            ),
            (
                "Investigate returns drivers in "
                f"{regional['highest_returns_region']}."
            ),
            (
                "Build a recovery plan for "
                f"{regional['weakest_growth_region']} "
                "while protecting momentum in "
                f"{regional['strongest_growth_region']}."
            ),
        ],

        "closing_statement": (
            "Protect growth by converting the clearest "
            "customer and operational exceptions into "
            "named actions for the next quarter."
        ),
    }


def generate_executive_insights(
    snapshot: dict[str, Any],
    *,
    allow_fallback: bool = False,
) -> dict[str, Any]:
    """
    Run ADK on deterministic facts and return
    presentation-ready content.
    """

    configure_business_snapshot(
        snapshot
    )

    try:
        return asyncio.run(
            _generate_with_adk()
        )
    except Exception:
        if not allow_fallback:
            raise

        return _deterministic_fallback(
            snapshot
        )


if __name__ == "__main__":
    print(
        "NovaRetail ADK agent loaded. "
        "Run `adk web .` from the project root "
        "or `python -m NovaRetail.main` for the "
        "automated deck."
    )