"""
Router — classifies user requests into operation types.
"""

import json
import re
from typing import Optional

from langchain_agent import llm
from langchain_core.messages import HumanMessage, SystemMessage

ROUTER_PROMPT = """You route Google Slides requests. Given a brief slide summary and user message, output ONE JSON object:
{"operation": "<type>", "message": "<brief description of what to do>"}

Operation types:
- "edit_layout": move, resize, reposition, align, center, or make symmetrical EXISTING elements
- "create_content": add new shapes, text boxes, rectangles, lines, or sections to the CURRENT slide
- "create_slide": create a BRAND NEW slide (ONLY when user says "new slide", "create a slide", "add a slide")
- "edit_text": change text content, font, size, color, bold/italic of EXISTING elements
- "answer_question": answer a question about the presentation (fonts, content, structure, etc.)

Rules:
- "add a section", "add a box", "add text about X", "draw a line" → create_content (NOT create_slide)
- "make this symmetrical", "center these", "move X to the right" → edit_layout
- "change the font to Arial", "make text bigger", "replace the title text", "change the background/fill color", "match the colors", "change colors of text boxes" → edit_text
- "what font is on slide 2?", "what does slide 5 say?" → answer_question
- If the current slide has NO free space and the user wants to add substantial content, use "create_slide"
- EMPTY TEXT BOXES on slide (see context): If the context lists empty text boxes and the user wants wording inside a box (e.g. conclusion, sentences, paragraph, bullets, "fill it", "textbox with ...", "put the text in the box") WITHOUT clearly asking for an ADDITIONAL separate box, use "edit_text" so the model can replace_text on an existing objectId. Do NOT use "create_content" for that — it stacks a duplicate box on top.
- If the user clearly wants ANOTHER / NEW / SECOND / DUPLICATE text box in addition to existing ones (e.g. "add a new empty text box", "another text box on the right"), use "create_content".
- Output ONLY the JSON object, no markdown."""


def build_router_context(
    total_slides: int, title: str, current_index: Optional[int],
    page_w_pt: float, page_h_pt: float, num_elements: int,
    gaps: list[tuple[float, float, float]],
    empty_text_boxes_hint: Optional[str] = None,
) -> str:
    """Build a short context string for the router (no full element list)."""
    lines = [
        f'Presentation: "{title}" ({total_slides} slides)',
        f"Current slide: index {current_index}, {page_w_pt}x{page_h_pt} PT, {num_elements} elements",
    ]
    if gaps:
        gap_strs = [f"y={g[0]}-{g[1]} ({g[2]}pt)" for g in gaps]
        lines.append(f"Free vertical space: {', '.join(gap_strs)}")
    else:
        lines.append("Free vertical space: NONE (slide is full)")
    if empty_text_boxes_hint:
        lines.append(empty_text_boxes_hint)
    return "\n".join(lines)


def route_request(context_summary: str, user_message: str) -> tuple[str, str]:
    """Route user request to an operation type. Returns (operation, message)."""
    messages = [
        SystemMessage(content=ROUTER_PROMPT),
        HumanMessage(content=f"{context_summary}\n\nUser request: {user_message}"),
    ]
    response = llm.invoke(messages)
    text = response.content if hasattr(response, "content") else str(response)
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()
    print(f"   ROUTER raw: {text[:300]}")
    try:
        parsed = json.loads(text)
        return parsed.get("operation", "edit_layout"), parsed.get("message", "")
    except json.JSONDecodeError:
        return "edit_layout", ""
