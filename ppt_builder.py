from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt


# ============================================================
# COLORS
# ============================================================

WHITE = RGBColor(255, 255, 255)
LIGHT_TEXT = RGBColor(230, 234, 245)
MUTED_TEXT = RGBColor(190, 200, 225)
DARK_TEXT = RGBColor(25, 31, 55)


# ============================================================
# FIND SHAPE
# ============================================================

def _shape_names(
    shape_name: str | Iterable[str],
) -> tuple[str, ...]:
    """Normalize one shape name or ordered fallback names."""

    if isinstance(shape_name, str):
        names = (shape_name,)
    else:
        try:
            names = tuple(shape_name)
        except TypeError as exc:
            raise TypeError(
                "shape_name must be a string or an iterable of strings."
            ) from exc

    if not names or any(
        not isinstance(name, str) or not name.strip()
        for name in names
    ):
        raise ValueError(
            "Shape names must be non-empty strings."
        )

    return names


def find_shape(
    slide,
    shape_name: str | Iterable[str],
):
    """Find a PowerPoint shape by one name or fallback names."""

    shape_names = _shape_names(shape_name)

    # Respect the caller's fallback order during exact matching.
    shapes_by_name = {
        shape.name: shape
        for shape in slide.shapes
    }

    for name in shape_names:
        if name in shapes_by_name:
            return shapes_by_name[name]

    # Case-insensitive fallback while preserving fallback order.
    shapes_by_normalized_name = {
        shape.name.strip().casefold(): shape
        for shape in slide.shapes
    }

    for name in shape_names:
        normalized_name = name.strip().casefold()

        if normalized_name in shapes_by_normalized_name:
            return shapes_by_normalized_name[
                normalized_name
            ]

    available_shapes = [
        shape.name
        for shape in slide.shapes
    ]

    raise KeyError(
        f"None of the shapes {shape_names!r} was found. "
        f"Available shapes: {available_shapes}"
    )


# ============================================================
# REMOVE SHAPE
# ============================================================

def remove_shape(shape):
    shape_element = shape._element
    shape_element.getparent().remove(
        shape_element
    )


# ============================================================
# REPLACE PLACEHOLDER WITH IMAGE
# ============================================================

def replace_placeholder_with_image(
    slide,
    shape_name: str | Iterable[str],
    image_path,
):
    placeholder = find_shape(
        slide,
        shape_name,
    )

    image_path = Path(image_path)

    if not image_path.is_file():
        raise FileNotFoundError(
            f"Image was not found:\n{image_path}"
        )

    left = placeholder.left
    top = placeholder.top
    width = placeholder.width
    height = placeholder.height

    # Capture the resolved string name before removing the shape.
    placeholder_name = placeholder.name

    remove_shape(
        placeholder
    )

    picture = slide.shapes.add_picture(
        str(image_path),
        left,
        top,
        width,
        height,
    )

    # This uses the actual resolved placeholder name rather than
    # converting a tuple of fallback names into a shape name.
    picture.name = (
        f"{placeholder_name}_generated"
    )

    return picture


# ============================================================
# REPLACE ICON
# ============================================================

def replace_icon(
    slide,
    shape_name: str | Iterable[str],
    icon_path,
):
    return replace_placeholder_with_image(
        slide=slide,
        shape_name=shape_name,
        image_path=icon_path,
    )


def replace_icon_if_available(
    slide,
    shape_name: str | Iterable[str],
    icon_path,
):
    """
    Replace an optional icon when both its file and one of its
    possible placeholders exist.
    """

    icon_path = Path(icon_path)

    if not icon_path.is_file():
        print(
            f"Skipping missing icon: {icon_path}"
        )
        return None

    try:
        icon_placeholder = find_shape(
            slide,
            shape_name,
        )
    except KeyError:
        print(
            "Skipping missing icon placeholder: "
            f"{_shape_names(shape_name)}"
        )
        return None

    return replace_icon(
        slide=slide,
        shape_name=icon_placeholder.name,
        icon_path=icon_path,
    )


# ============================================================
# DETERMINE TEXT STYLE
# ============================================================

def get_text_style(
    shape_name: str,
):
    """
    Return the font size, bold setting, color, and alignment for
    a resolved PowerPoint shape.
    """

    normalized_name = (
        shape_name.strip().casefold()
    )

    if normalized_name in {
        "kpi1_value",
        "kpi2_value",
        "kpi3_value",
        "kpi_1_value",
        "kpi_2_value",
        "kpi_3_value",
    }:
        return (
            26,
            True,
            WHITE,
            PP_ALIGN.CENTER,
        )

    if normalized_name in {
        "kpi1_label",
        "kpi2_label",
        "kpi3_label",
        "kpi_1_label",
        "kpi_2_label",
        "kpi_3_label",
    }:
        return (
            13,
            False,
            LIGHT_TEXT,
            PP_ALIGN.CENTER,
        )

    if normalized_name in {
        "subtitle_1",
        "subtitle1",
        "subtitle 1",
    }:
        return (
            16,
            False,
            LIGHT_TEXT,
            PP_ALIGN.LEFT,
        )

    if normalized_name in {
        "summary_text",
        "executive_summary",
        "executive_summary_text",
    }:
        return (
            15,
            False,
            LIGHT_TEXT,
            PP_ALIGN.LEFT,
        )

    if normalized_name in {
        "closing_statement",
        "closing_text",
        "takeaway_text",
    }:
        return (
            15,
            True,
            LIGHT_TEXT,
            PP_ALIGN.LEFT,
        )

    if normalized_name in {
        "title_2",
        "performance_title",
        "title_performance",
        "title_3",
        "customer_title",
        "title_customer",
        "title_4",
        "risk_title",
        "title_risk",
    }:
        return (
            22,
            True,
            DARK_TEXT,
            PP_ALIGN.LEFT,
        )

    if normalized_name in {
        "title_5",
        "recommendations_title",
        "actions_title",
        "title_recommendations",
    }:
        return (
            22,
            True,
            WHITE,
            PP_ALIGN.LEFT,
        )

    return (
        None,
        None,
        None,
        None,
    )


# ============================================================
# SET TEXT
# ============================================================

def set_text(
    slide,
    shape_name: str | Iterable[str],
    text,
    font_size=None,
    bold=None,
    color=None,
    alignment=None,
):
    shape = find_shape(
        slide,
        shape_name,
    )

    if not shape.has_text_frame:
        raise ValueError(
            f"{shape.name!r} does not contain a text frame."
        )

    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    (
        default_size,
        default_bold,
        default_color,
        default_alignment,
    ) = get_text_style(
        shape.name
    )

    if font_size is None:
        font_size = default_size

    if bold is None:
        bold = default_bold

    if color is None:
        color = default_color

    if alignment is None:
        alignment = default_alignment

    paragraph = text_frame.paragraphs[0]
    paragraph.text = str(text)

    if alignment is not None:
        paragraph.alignment = alignment

    for run in paragraph.runs:
        if font_size is not None:
            run.font.size = Pt(
                font_size
            )

        if bold is not None:
            run.font.bold = bold

        if color is not None:
            if isinstance(
                color,
                RGBColor,
            ):
                run.font.color.rgb = color
            else:
                run.font.color.rgb = RGBColor(
                    *color
                )

    return shape


# ============================================================
# SET BULLETS
# ============================================================

def _clean_bullets(
    bullets,
) -> list[str]:
    if bullets is Ellipsis:
        raise ValueError(
            "Bullets cannot be Ellipsis. "
            "Pass the corresponding insights list."
        )

    if isinstance(
        bullets,
        (str, bytes),
    ):
        raise TypeError(
            "bullets must be an iterable of bullet values, "
            "not a string."
        )

    try:
        raw_bullets = list(
            bullets
        )
    except TypeError as exc:
        raise TypeError(
            "bullets must be an iterable of bullet values."
        ) from exc

    cleaned: list[str] = []

    for bullet in raw_bullets:
        if bullet is Ellipsis:
            raise ValueError(
                "A bullet value cannot be Ellipsis. "
                "Pass generated insight text."
            )

        if bullet is None:
            continue

        text = str(
            bullet
        ).strip()

        if text:
            cleaned.append(
                text
            )

    if not cleaned:
        raise ValueError(
            "At least one non-empty bullet is required."
        )

    return cleaned


def set_bullets(
    slide,
    shape_name: str | Iterable[str],
    bullets,
    font_size=14,
    color=None,
):
    shape = find_shape(
        slide,
        shape_name,
    )

    if not shape.has_text_frame:
        raise ValueError(
            f"{shape.name!r} does not have a text frame."
        )

    cleaned_bullets = _clean_bullets(
        bullets
    )

    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    if color is None:
        if shape.name.strip().casefold() in {
            "recommendations_text",
            "actions_text",
            "recommendation_text",
            "insights_recommendations",
            "body_5",
        }:
            text_color = LIGHT_TEXT
        else:
            text_color = DARK_TEXT

    elif isinstance(
        color,
        RGBColor,
    ):
        text_color = color

    else:
        text_color = RGBColor(
            *color
        )

    for index, bullet in enumerate(
        cleaned_bullets
    ):
        if index == 0:
            paragraph = (
                text_frame.paragraphs[0]
            )
        else:
            paragraph = (
                text_frame.add_paragraph()
            )

        if bullet.startswith("•"):
            paragraph.text = bullet
        else:
            paragraph.text = (
                f"• {bullet}"
            )

        paragraph.level = 0
        paragraph.space_after = Pt(8)
        paragraph.line_spacing = 1.05

        for run in paragraph.runs:
            run.font.size = Pt(
                font_size
            )
            run.font.color.rgb = (
                text_color
            )

    return shape


# ============================================================
# OPTIONAL: FORMAT STANDARD TEXT BOX
# ============================================================

def format_text_box(
    slide,
    shape_name: str | Iterable[str],
    font_size=14,
    bold=False,
    color=(255, 255, 255),
):
    shape = find_shape(
        slide,
        shape_name,
    )

    if not shape.has_text_frame:
        return shape

    text_frame = shape.text_frame
    text_frame.word_wrap = True

    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(
                font_size
            )
            run.font.bold = bold
            run.font.color.rgb = RGBColor(
                *color
            )

    return shape


def remove_shape_if_exists(
    slide,
    shape_name: str | Iterable[str],
):
    try:
        shape = find_shape(
            slide,
            shape_name,
        )
    except KeyError:
        return False

    remove_shape(
        shape
    )

    return True


def remove_template_instructions(
    slide,
):
    """
    Remove instructional/helper text boxes left in the
    PowerPoint template after charts and icons are populated.
    """

    shapes_to_remove = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        text = shape.text.strip()

        if not text:
            continue

        text_lower = text.lower()
        shape_name_lower = shape.name.strip().lower()

        is_chart_instruction = (
            text_lower == "chart"
            or "insert your chart here" in text_lower
            or "suggested:" in text_lower
        )

        is_icon_instruction = (
            text_lower == "icon"
            or "pick the icon" in text_lower
            or (
                shape_name_lower.startswith("icon_")
                and shape_name_lower.endswith("_label")
            )
        )

        if (
            is_chart_instruction
            or is_icon_instruction
        ):
            shapes_to_remove.append(shape)

    for shape in shapes_to_remove:
        print(
            "Removing template instruction: "
            f"{shape.name!r} -> {shape.text!r}"
        )

        remove_shape(shape)