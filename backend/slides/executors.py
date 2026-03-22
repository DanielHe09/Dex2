"""
Executor prompts and LLM call helper — one focused prompt per operation type.
"""

import json
import re

from langchain_agent import llm
from langchain_core.messages import HumanMessage, SystemMessage


# ---------------------------------------------------------------------------
EDIT_LAYOUT_PROMPT = """You reposition and resize elements on a Google Slides slide. Output a JSON object:
{"instructions": [...], "message": "brief summary"}

Each instruction: {"action": "move"|"resize"|"move_and_resize", "objectId": "...", "x_pt": ..., "y_pt": ..., "width_pt": ..., "height_pt": ...}
- "move": set x_pt and/or y_pt
- "resize": set width_pt and/or height_pt
- "move_and_resize": set any combination

Rules:
- Positions from top-left (0,0). Units: PT (points).
- Only include elements you want to change.
- Slide center X = half slide width, center Y = half slide height.
- For symmetry: mirror positions relative to slide center, account for element sizes.
- No markdown fences. Output ONLY the JSON object."""


# ---------------------------------------------------------------------------
CREATE_CONTENT_PROMPT = """You add new shapes, text boxes, and lines to a Google Slides slide. Output a JSON object:
{"instructions": [...], "message": "brief summary"}

Shape instruction (with coordinates):
{"action": "create_shape", "shape_type": "TEXT_BOX"|"RECTANGLE"|"ROUND_RECTANGLE"|"ELLIPSE"|"TRIANGLE"|etc., "x_pt": ..., "y_pt": ..., "width_pt": ..., "height_pt": ..., "text": "...", "background_color": "#hex", "border_color": "#hex", "border_weight_pt": ..., "font_size_pt": ..., "bold": true/false, "italic": true/false, "font_family": "...", "color": "#hex"}

For NUMBERED LISTS or "N things" layouts (e.g. "2 things I hated", "3 things I learned"), use ROLE-based instructions so layout is automatic. Omit x_pt, y_pt, width_pt, height_pt and use "role" instead:
- {"action": "create_shape", "role": "title", "shape_type": "TEXT_BOX", "text": "...", ...style...}
- {"action": "create_shape", "role": "item_1_number", "shape_type": "TEXT_BOX", "text": "1", ...style...}
- {"action": "create_shape", "role": "item_1_text", "shape_type": "TEXT_BOX", "text": "First point here", ...style...}
- {"action": "create_shape", "role": "item_2_number", "shape_type": "TEXT_BOX", "text": "2", ...style...}
- {"action": "create_shape", "role": "item_2_text", "shape_type": "TEXT_BOX", "text": "Second point", ...style...}
Use roles "title", "item_N_number", "item_N_text" for each numbered item. Positions will be assigned automatically and aligned.

Line instruction:
{"action": "create_line", "line_type": "STRAIGHT"|"BENT"|"CURVED", "start_x_pt": ..., "start_y_pt": ..., "end_x_pt": ..., "end_y_pt": ..., "color": "#hex", "weight_pt": ...}

MUST GET RIGHT (use exact values from "Presentation visual style" and element list):
1. Text font: Set "font_family" to one of the Common fonts from the presentation style for EVERY text shape. No other fonts.
2. Text color: Set "color" to one of the Common text colors from the presentation style (often #000000). Use the exact hex from the deck.
3. Text box fill: Set "background_color" to the exact hex from Common background fills that matches similar elements (e.g. title box fill = same as other title boxes; list/body box fill = same as other list boxes). Use values from the element list (bg:) when given.
4. Text box outline: Set "border_color" from Common border/outline colors or from element list (border: #hex). Set "border_weight_pt" (e.g. 1–2) when existing slides show a visible outline (see element list border_weight). If boxes on existing slides have a border, copy that.

Rules:
- If the slide description includes *** EMPTY TEXT BOXES *** and the user wants text inside a box (conclusion, sentences, fill the box, "textbox with ...") without clearly asking for an ADDITIONAL new box, use replace_text with the correct objectId — do NOT create_shape a duplicate TEXT_BOX on top.
- ALWAYS include "text" with actual content when user asks to add a section/content about something (unless you are only using replace_text to fill an empty box as above).
- When the user asks to add a section/column/card "in the empty space", "next to 2", or "that matches the others", follow the "Current slide visual style" section/column instructions exactly: create an outer RECTANGLE (section_background_fill, section_border_color), then a number label, then a title TEXT_BOX (title_text_color, bold), then a body TEXT_BOX (inner_text_box_fill, inner_text_box_border, body text color). Place and size the new block to align with existing columns and fill the gap.
- For numbered lists or "N things" / "N items" slides, PREFER role-based create_shape (title, item_1_number, item_1_text, item_2_number, item_2_text, ...) so number and text boxes align and do not overlap.
- Use ONLY hex values from the presentation style and element list for background_color, color, and border_color. Use shape_type TEXT_BOX or RECTANGLE (or ROUND_RECTANGLE if the deck uses rounded corners).
- Use RECTANGLE or ROUND_RECTANGLE with background_color for card-style sections.
- When using coordinates (no role), a new element at y=Y with height=H MUST fit within a free-space gap: Y >= gap_start AND Y + H <= gap_end.
- If content doesn't fit, SHRINK height or font size to fit within the gap.
- **Empty TEXT_BOX in whitespace under the title** (above a row of cards/columns): **Match the row below** — copy **border_color**, **border_weight_pt**, and **background_color** from the **Elements** list for those body/card boxes (same grey outline as the feasibility columns), not a black title-accent border. **Align x_pt and width_pt** with the **combined span of that column row** (left edge of the first column through right edge of the last), not wider than the slide. **height_pt** ≥ **72pt** (or ~65–85% of the FREE SPACE gap); never a sliver under 48pt tall. Leave **at least ~12–24pt vertical space** between the bottom of this box and the top of the boxes/columns below (shrink height_pt or raise y_pt slightly if needed) so nothing touches.
- You can also move/resize existing elements to make room: {"action": "move"|"resize"|"move_and_resize", "objectId": "...", ...}
- No markdown fences. Output ONLY the JSON object."""


# ---------------------------------------------------------------------------
CREATE_SLIDE_PROMPT = """You design a new Google Slides slide from scratch using shapes. The slide will be created as BLANK, then you populate it with shapes that match the presentation's visual style.

Output a JSON object:
{"instructions": [...], "message": "brief summary", "insert_after": "current"|"end"|<index>}

Each instruction creates a shape on the NEW slide. You may use EITHER coordinates OR roles:

With coordinates:
{"action": "create_shape", "shape_type": "TEXT_BOX"|"RECTANGLE"|"ROUND_RECTANGLE"|etc., "x_pt": ..., "y_pt": ..., "width_pt": ..., "height_pt": ..., "text": "...", "background_color": "#hex", "border_color": "#hex", "border_weight_pt": ..., "font_size_pt": ..., "bold": true/false, "italic": true/false, "font_family": "...", "color": "#hex"}

For NUMBERED LISTS or "N things" slides (e.g. "2 things I hated", "3 new things I learned"), use ROLES and omit x_pt, y_pt, width_pt, height_pt so layout is automatic:
- {"action": "create_shape", "role": "title", "shape_type": "TEXT_BOX", "text": "2 things I hated", ...style...}
- {"action": "create_shape", "role": "item_1_number", "shape_type": "TEXT_BOX", "text": "1", ...style...}
- {"action": "create_shape", "role": "item_1_text", "shape_type": "TEXT_BOX", "text": "First point", ...style...}
- {"action": "create_shape", "role": "item_2_number", "shape_type": "TEXT_BOX", "text": "2", ...style...}
- {"action": "create_shape", "role": "item_2_text", "shape_type": "TEXT_BOX", "text": "Second point", ...style...}
Use roles "title", "item_N_number", "item_N_text" for each item. Positions will be assigned and aligned automatically.

MUST GET RIGHT (use exact values from "Presentation visual style" and element list):
1. Text font: Set "font_family" to one of the Common fonts for EVERY text shape. Use "bold": true for title-style text when the deck does. No other fonts.
2. Text color: Set "color" to one of the Common text colors (often #000000). Exact hex from the deck.
3. Text box fill: Set "background_color" to the exact hex from Common background fills that matches similar elements (title box = same fill as other title boxes; list/body boxes = same fill as other list boxes). Use element list (bg:) when given.
4. Text box outline: Set "border_color" from Common border/outline colors or from element list (border: #hex). Set "border_weight_pt" from element list (border_weight) when present, or use 1–2 when existing slides show a visible outline. Copy border style from existing slides.

Rules:
- For "N things", "N items", numbered lists, PREFER role-based instructions (title, item_1_number, item_1_text, ...) so number and text boxes align and do not overlap.
- Design a COMPLETE slide layout — add title, body content, decorative elements as needed.
- Use ONLY hex values from the presentation style and element list. Use shape_type TEXT_BOX, RECTANGLE, or ROUND_RECTANGLE to match existing box shapes (sharp vs rounded corners).
- Cover the full slide area (dimensions provided). Don't leave the slide empty.
- Generate real, substantive content for the topic the user requested.
- Use shapes with background_color for card-style sections, matching existing bg: colors.

COPY FROM CURRENT SLIDE when the user asks to copy the heading, footer, or table of contents:
- HEADING: Look at the CURRENT SLIDE elements at the top. If there is a vertical line + title + subtitle layout, REPLICATE it: create a thin vertical RECTANGLE (same color as original), then a title shape, then a subtitle shape — same relative positions (x_pt, y_pt) and sizes (width_pt, height_pt), same fonts and colors; only change the title/subtitle text to fit the new slide topic.
- TABLE OF CONTENTS / FOOTER: Look at the CURRENT SLIDE elements at the bottom. If there are multiple separate shapes (tabs like "Business Overview", "Market Feasibility", etc.), REPLICATE that structure: create one create_shape per tab, same count, same relative positions and sizes, same bg: and style; you may update the tab labels to match the new slide (e.g. keep "Conclusion" or add "Long Term Trends") but keep the same visual layout.
- Do not replace the heading or ToC with a single simplified shape — copy the actual structure (multiple elements) so it looks like the original.

- No markdown fences. Output ONLY the JSON object."""


# ---------------------------------------------------------------------------
EDIT_TEXT_PROMPT = """You edit text content, formatting, and shape fill/outline on existing Google Slides elements. Output a JSON object:
{"instructions": [...], "message": "brief summary"}

Replace text: {"action": "replace_text", "objectId": "...", "new_text": "..."}
Update text style: {"action": "update_text_style", "objectId": "...", "font_size_pt": ..., "bold": true/false, "italic": true/false, "underline": true/false, "font_family": "...", "color": "#hex"}
Update shape fill/background (for text boxes and shapes): {"action": "update_shape_fill", "objectId": "...", "background_color": "#hex", "border_color": "#hex", "border_weight_pt": number (optional, omit or use 0 for no border)}

Rules:
- If *** EMPTY TEXT BOXES *** are listed and the user wants body copy, a conclusion, bullets, or sentences in "the" / "that" text box, use replace_text on one of those objectIds (pick the empty box that matches the user's intent by position/size). Do not create_shape a new text box on top of an empty one.
- replace_text replaces ALL text in the element.
- update_text_style applies to ALL text in the element (color = text color).
- update_shape_fill changes the shape's background fill and/or border. Use when the user says "change the background/fill color", "match the colors", "make these boxes the same color as the rest of the slideshow". Use hex values from the slide's existing elements (style and bg: in the element list) or from "Common background fills" / "Common text colors" in the presentation style.
- You can combine replace_text, update_text_style, and update_shape_fill on the same or different elements.
- When the user wants colors to match the slideshow, apply update_shape_fill (and update_text_style for text color) to the relevant elements using the deck's actual colors from the context.
- No markdown fences. Output ONLY the JSON object."""


# ---------------------------------------------------------------------------
ANSWER_QUESTION_PROMPT = """You answer questions about a Google Slides presentation. Output a JSON object:
{"instructions": [], "message": "your detailed answer"}

Rules:
- Always return an empty instructions array.
- Put your answer in "message".
- Use the slide data provided to give accurate, specific answers about fonts, colors, content, structure, etc.
- No markdown fences. Output ONLY the JSON object."""


# ---------------------------------------------------------------------------
EXECUTOR_PROMPTS = {
    "edit_layout": EDIT_LAYOUT_PROMPT,
    "create_content": CREATE_CONTENT_PROMPT,
    "create_slide": CREATE_SLIDE_PROMPT,
    "edit_text": EDIT_TEXT_PROMPT,
    "answer_question": ANSWER_QUESTION_PROMPT,
}


def call_executor(
    system_prompt: str, context: str, user_message: str
) -> tuple[list[dict], str, dict]:
    """Call an executor LLM with a focused prompt. Returns (instructions, message, full_parsed_dict)."""
    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=f"{context}\n\nUser request: {user_message}"),
    ]
    response = llm.invoke(messages)
    text = response.content if hasattr(response, "content") else str(response)
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()
    print(f"   EXECUTOR raw: {text[:500]}")
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            instructions = parsed.get("instructions", [])
            if not isinstance(instructions, list):
                instructions = []
            return instructions, parsed.get("message", ""), parsed
        if isinstance(parsed, list):
            return parsed, "", {}
        return [], "", {}
    except json.JSONDecodeError as e:
        print(f"   EXECUTOR JSON parse error: {e}")
        return [], "", {}
