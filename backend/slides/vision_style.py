"""
Extract slide style (colors, font) from a slide screenshot using Gemini vision.
Used when API-based style (get_presentation_style_values) is missing or unreliable.
"""

import base64
import json
import os
import re
from typing import Any, Optional

import google.generativeai as genai


STYLE_PROMPT = """Look at this screenshot of a single slide from a presentation.

Extract the visual style so new content can match existing layout and formatting. Reply with a single JSON object (no markdown, no code fence).

REQUIRED keys (always include):
- "primary_text_color": body/content text color as hex (e.g. "#000000", "#ffffff")
- "primary_font": main font family (e.g. "Roboto", "Arial")
- "primary_background_fills": list of 1–3 hex colors for shape fills. Include the fill used for LARGE section/card/column blocks first if present (e.g. light blue for column backgrounds), then inner box fill (e.g. white) if different.
- "primary_border_colors": list of 1–2 hex colors for shape borders (e.g. ["#ffffff"] for white borders on columns)

OPTIONAL keys (include when the slide has a clear layout with sections/columns/cards):
- "section_background_fill": hex fill of large section/column/card blocks (the big blocks that contain title + content). Same as slide background if columns match it.
- "section_border_color": hex border color of those section blocks (e.g. "#ffffff" for white outline).
- "inner_text_box_fill": hex fill of smaller text boxes INSIDE sections (e.g. "#ffffff" for white content boxes).
- "inner_text_box_border": hex border of those inner boxes (e.g. "#ffffff" or "#e0e0e0").
- "title_text_color": hex color of section/column titles/headings (often same as an accent; e.g. blue).
- "title_bold": true if section titles are bold.

Describe what you see: if there are numbered columns or cards (e.g. "1", "2") with a distinct outer block and an inner content area, set section_* and inner_* so a new column "3" can be added with the same formatting. If the slide is simple (single text boxes), omit the optional keys.
Output only the JSON object, nothing else."""


def format_style_for_prompt(style_values: dict[str, Any]) -> str:
    """Format vision-extracted style for inclusion in executor context so the LLM uses these exact values."""
    font = style_values.get("primary_font") or "Arial"
    text_color = style_values.get("primary_text_color") or "#333333"
    fills = style_values.get("primary_background_fills") or ["#ffffff"]
    borders = style_values.get("primary_border_colors") or ["#000000"]
    fills_str = ", ".join(fills[:3]) if isinstance(fills, list) else str(fills)
    borders_str = ", ".join(borders[:2]) if isinstance(borders, list) else str(borders)
    lines = [
        "Current slide visual style (from image — use these EXACT values):",
        f"- font_family: {font}",
        f"- body text color (color): {text_color}",
        f"- background fill (background_color): {fills_str}",
        f"- border/outline (border_color): {borders_str}",
    ]
    # When slide has section/column layout, tell executor how to replicate a matching block
    section_fill = style_values.get("section_background_fill")
    section_border = style_values.get("section_border_color")
    inner_fill = style_values.get("inner_text_box_fill")
    inner_border = style_values.get("inner_text_box_border")
    title_color = style_values.get("title_text_color")
    title_bold = style_values.get("title_bold", True)
    if section_fill or inner_fill:
        lines.append("")
        lines.append("When the user asks to add a section/column/card in empty space that matches existing ones (e.g. 'add a column next to 2', 'slide into that empty space'):")
        lines.append("1. Create an outer RECTANGLE for the section: background_color = " + (section_fill or fills[0] if fills else "#f0f8ff") + ", border_color = " + (section_border or borders[0] if borders else "#ffffff") + ", border_weight_pt = 1.")
        lines.append("2. Inside it, add a small number label (e.g. '3') in the top-left: use section fill and border, or match existing number style.")
        lines.append("3. Add the section TITLE as a TEXT_BOX: font_family = " + font + ", color = " + (title_color or text_color) + ", bold = " + str(title_bold).lower() + ", background_color = " + (section_fill or fills[0] if fills else "#f0f8ff") + " (same as section), border_color = " + (section_border or borders[0] if borders else "#ffffff") + ".")
        lines.append("4. Add the body/content as a TEXT_BOX: background_color = " + (inner_fill or "#ffffff") + ", border_color = " + (inner_border or (borders[0] if borders else "#e0e0e0")) + ", color = " + text_color + ", font_family = " + font + ". Match size and position to existing columns.")
        lines.append("Use the same order, alignment, and spacing as the existing sections so the new block fits into the layout.")
    else:
        lines.append("Set font_family, color, background_color, and border_color to these values in every create_shape instruction.")
    return "\n".join(lines)


def _decode_screenshot_data(screenshot_base64: str) -> tuple[bytes, str]:
    """Return (image_bytes, mime_type). Handles raw base64 or data URL."""
    data = (screenshot_base64 or "").strip()
    if data.startswith("data:"):
        match = re.match(r"data:([^;]+);base64,(.+)", data, re.DOTALL)
        if match:
            mime = match.group(1).strip().lower()
            if "png" in mime:
                mime = "image/png"
            elif "jpeg" in mime or "jpg" in mime:
                mime = "image/jpeg"
            else:
                mime = "image/png"
            return base64.b64decode(match.group(2)), mime
        data = data.split(",", 1)[-1]
    return base64.b64decode(data), "image/png"


def extract_style_from_slide_image(screenshot_base64: Optional[str]) -> Optional[dict[str, Any]]:
    """
    Use Gemini vision to infer primary text color, font, background fills, and border colors
    from a slide screenshot. Returns a dict compatible with normalize_instructions_style:
    primary_font, primary_text_color, primary_background_fills, primary_border_colors.
    Returns None on parse/API error.
    """
    if not screenshot_base64 or not screenshot_base64.strip():
        return None
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        print("   VISION_STYLE: GOOGLE_API_KEY not set, skipping vision style")
        return None
    try:
        image_bytes, mime_type = _decode_screenshot_data(screenshot_base64)
    except Exception as e:
        print(f"   VISION_STYLE: failed to decode screenshot: {e}")
        return None
    genai.configure(api_key=api_key)
    # Use stable vision-capable model (gemini-1.5-flash is deprecated/removed)
    model = genai.GenerativeModel("gemini-2.5-flash")
    parts = [
        {"inline_data": {"mime_type": mime_type, "data": image_bytes}},
        STYLE_PROMPT,
    ]
    try:
        response = model.generate_content(parts)
        text = (response.text or "").strip()
    except Exception as e:
        print(f"   VISION_STYLE: Gemini API error: {e}")
        return None
    if not text:
        return None
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```\s*$", "", text)
    text = text.strip()
    try:
        out = json.loads(text)
    except json.JSONDecodeError as e:
        print(f"   VISION_STYLE: JSON parse error: {e}, raw={text[:200]!r}")
        return None
    primary_font = out.get("primary_font") or "Arial"
    primary_text_color = out.get("primary_text_color") or "#333333"
    primary_background_fills = out.get("primary_background_fills")
    if not isinstance(primary_background_fills, list):
        primary_background_fills = [primary_text_color] if primary_text_color else ["#ffffff"]
    primary_border_colors = out.get("primary_border_colors")
    if not isinstance(primary_border_colors, list):
        primary_border_colors = ["#000000"]
    result = {
        "primary_font": primary_font,
        "primary_text_color": primary_text_color,
        "primary_background_fills": primary_background_fills[:6],
        "primary_border_colors": primary_border_colors[:4],
    }
    # Optional layout-style keys for section/column/card matching
    if out.get("section_background_fill"):
        result["section_background_fill"] = out.get("section_background_fill")
    if out.get("section_border_color"):
        result["section_border_color"] = out.get("section_border_color")
    if out.get("inner_text_box_fill"):
        result["inner_text_box_fill"] = out.get("inner_text_box_fill")
    if out.get("inner_text_box_border"):
        result["inner_text_box_border"] = out.get("inner_text_box_border")
    if out.get("title_text_color"):
        result["title_text_color"] = out.get("title_text_color")
    if "title_bold" in out:
        result["title_bold"] = bool(out.get("title_bold"))
    print(
        f"   VISION_STYLE: primary_font={primary_font!r} primary_text_color={primary_text_color!r} "
        f"primary_background_fills={primary_background_fills!r} primary_border_colors={primary_border_colors!r}"
    )
    if result.get("section_background_fill"):
        print(
            f"   VISION_STYLE: section_fill={result['section_background_fill']!r} section_border={result.get('section_border_color')!r} "
            f"inner_fill={result.get('inner_text_box_fill')!r} title_color={result.get('title_text_color')!r}"
        )
    return result


CREATE_CONTENT_VISION_PROMPT_TEMPLATE = """You are looking at a screenshot of a single slide from a Google Slides presentation.

Slide dimensions: {page_width_pt} x {page_height_pt} points (origin top-left, coordinates in points).

User request: {user_message}
{layout_context_block}

Your task: Output a JSON object with instructions to ADD the requested content so it matches the slide's existing style and layout. Use this exact format (no markdown, no code fence):
{{"instructions": [...], "message": "brief summary"}}

Instructions:
- If the layout context lists *** EMPTY TEXT BOXES *** and the user wants to put wording inside a box (conclusion, sentences, paragraph, fill the box, etc.) and is NOT clearly asking for an extra brand-new box, output ONLY replace_text instructions: {{"action": "replace_text", "objectId": "<exact id from layout>", "new_text": "..."}}. Optionally add update_text_style on the same objectId to match font size/color. Do NOT create_shape a new TEXT_BOX on top — that stacks duplicates.
- Each new element must otherwise be a create_shape: {{"action": "create_shape", "shape_type": "TEXT_BOX" or "RECTANGLE", "x_pt": number, "y_pt": number, "width_pt": number, "height_pt": number, "text": "..." (for TEXT_BOX), "background_color": "#hex", "border_color": "#hex", "border_weight_pt": 1, "font_family": "...", "color": "#hex", "font_size_pt": number, "bold": true/false}}
- You MUST include x_pt, y_pt, width_pt, height_pt for every create_shape. Use FREE SPACE / GAP and Elements; for new columns, align with existing column positions.
- **Empty TEXT_BOX in the band under the title (above 3 column cards):** Match those **card/body** boxes — same **border_color** (usually light grey like the columns, **not** black), **border_weight_pt**, **background_color**, and align **x_pt / width_pt** with the **left–right span of the column row** (do not stretch full slide edge-to-edge). **height_pt** ≥ **72pt** (or ~70–85% of the GAP); no paper-thin strips. Keep **12–24pt clear gap** between this box’s bottom and the tops of the column cards (use Element y positions); shorten the box if it would touch.
- Match visual style of **the same class of element** (body cards vs title chrome).
- When adding a new section/column that should match existing ones (e.g. "3" next to "1" and "2"): create 1) an outer RECTANGLE for the section, 2) a small TEXT_BOX for the number label (e.g. "3"), 3) a TEXT_BOX for the section title, 4) a TEXT_BOX for the body text. Position them using the coordinates from the layout context so the new column sits in the empty space and aligns with existing columns.
- Do not create move/resize instructions for existing elements unless the user asked to move something. Only create new shapes when replace_text is not the right approach.
Output only the JSON object, nothing else."""


def generate_content_instructions_from_image(
    screenshot_base64: str,
    user_message: str,
    page_width_pt: float,
    page_height_pt: float,
    layout_context: Optional[str] = None,
) -> tuple[list[dict], str]:
    """
    Use Gemini vision to generate create_shape (and optionally create_line) instructions
    directly from the slide screenshot and user request. Returns (instructions, message).
    When layout_context is provided (current slide description with FREE SPACE gaps and
    element positions from the Slides API), Gemini can place new content in the correct
    empty region.
    """
    if not screenshot_base64 or not screenshot_base64.strip():
        return [], ""
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        print("   VISION_CONTENT: GOOGLE_API_KEY not set")
        return [], ""
    try:
        image_bytes, mime_type = _decode_screenshot_data(screenshot_base64)
    except Exception as e:
        print(f"   VISION_CONTENT: failed to decode screenshot: {e}")
        return [], ""
    layout_context_block = ""
    if layout_context and layout_context.strip():
        layout_context_block = "\n\nCurrent slide layout (from API — use FREE SPACE and Elements to place new content in the right spot):\n" + layout_context.strip()
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")
    prompt = CREATE_CONTENT_VISION_PROMPT_TEMPLATE.format(
        user_message=user_message,
        page_width_pt=page_width_pt,
        page_height_pt=page_height_pt,
        layout_context_block=layout_context_block,
    )
    parts = [
        {"inline_data": {"mime_type": mime_type, "data": image_bytes}},
        prompt,
    ]
    try:
        response = model.generate_content(parts)
        text = (response.text or "").strip()
    except Exception as e:
        print(f"   VISION_CONTENT: Gemini API error: {e}")
        return [], ""
    if not text:
        return [], ""
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```\s*$", "", text)
    text = text.strip()
    try:
        out = json.loads(text)
    except json.JSONDecodeError as e:
        print(f"   VISION_CONTENT: JSON parse error: {e}, raw={text[:300]!r}")
        return [], ""
    instructions = out.get("instructions")
    if not isinstance(instructions, list):
        instructions = []
    message = out.get("message") or ""
    print(f"   VISION_CONTENT: Gemini returned {len(instructions)} instructions")
    return instructions, message
