"""
Google Slides API module: read slide structure, use LLM to plan layout changes,
translate to Slides API batchUpdate requests.
"""

import json
import re
import uuid
from typing import Optional
import httpx
from langchain_agent import llm
from langchain_core.messages import HumanMessage, SystemMessage


SLIDES_API = "https://slides.googleapis.com/v1/presentations"

DEFAULT_PAGE_WIDTH_EMU = 9_144_000
DEFAULT_PAGE_HEIGHT_EMU = 6_858_000
PT_TO_EMU = 12_700


def parse_slides_url(url: str) -> Optional[tuple[str, Optional[str]]]:
    """Extract (presentation_id, page_object_id) from a Google Slides URL."""
    if not url or "docs.google.com/presentation" not in url:
        return None
    pres_match = re.search(r"/presentation/d/([a-zA-Z0-9_-]+)", url)
    if not pres_match:
        return None
    presentation_id = pres_match.group(1)
    page_match = re.search(r"#slide=id\.([a-zA-Z0-9_p.]+)", url)
    page_id = page_match.group(1) if page_match else None
    return (presentation_id, page_id)


def get_page_size(presentation_id: str, access_token: str) -> tuple[int, int]:
    """Return (width_emu, height_emu) for the presentation."""
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        r = httpx.get(
            f"{SLIDES_API}/{presentation_id}",
            headers=headers,
            params={"fields": "pageSize"},
            timeout=15.0,
        )
        r.raise_for_status()
        page_size = r.json().get("pageSize", {})
        w = page_size.get("width", {}).get("magnitude", DEFAULT_PAGE_WIDTH_EMU)
        h = page_size.get("height", {}).get("magnitude", DEFAULT_PAGE_HEIGHT_EMU)
        return (int(w), int(h))
    except Exception:
        return (DEFAULT_PAGE_WIDTH_EMU, DEFAULT_PAGE_HEIGHT_EMU)


def get_full_presentation(presentation_id: str, access_token: str) -> dict:
    """Fetch the entire presentation (all slides, elements, page size) in one API call."""
    headers = {"Authorization": f"Bearer {access_token}"}
    r = httpx.get(
        f"{SLIDES_API}/{presentation_id}",
        headers=headers,
        timeout=30.0,
    )
    r.raise_for_status()
    return r.json()


def get_slide_elements(
    presentation_id: str, page_id: str, access_token: str
) -> dict:
    """Fetch a single slide's page elements via presentations.pages.get."""
    headers = {"Authorization": f"Bearer {access_token}"}
    r = httpx.get(
        f"{SLIDES_API}/{presentation_id}/pages/{page_id}",
        headers=headers,
        timeout=15.0,
    )
    r.raise_for_status()
    return r.json()


def _summarize_element(el: dict) -> Optional[dict]:
    """
    Summarize a page element into a readable dict for the LLM.
    Returns None if the element has no transform (not movable).
    """
    transform = el.get("transform")
    size_obj = el.get("size", {})
    if not transform:
        return None

    w_dim = size_obj.get("width", {})
    h_dim = size_obj.get("height", {})
    w_mag = w_dim.get("magnitude", 0)
    h_mag = h_dim.get("magnitude", 0)
    w_unit = w_dim.get("unit", "EMU")
    h_unit = h_dim.get("unit", "EMU")

    # Convert to PT for readability
    width_pt = (w_mag / PT_TO_EMU) if w_unit == "EMU" else w_mag
    height_pt = (h_mag / PT_TO_EMU) if h_unit == "EMU" else h_mag

    scale_x = transform.get("scaleX", 1)
    scale_y = transform.get("scaleY", 1)
    tx = transform.get("translateX", 0)
    ty = transform.get("translateY", 0)

    # Actual rendered size
    rendered_w = width_pt * abs(scale_x)
    rendered_h = height_pt * abs(scale_y)

    # Position in PT
    x_pt = tx / PT_TO_EMU
    y_pt = ty / PT_TO_EMU

    # Determine element type, text content, and text style
    el_type = "shape"
    text_content = ""
    text_style_desc = ""
    if "shape" in el:
        shape_type = el["shape"].get("shapeType", "RECTANGLE")
        if shape_type == "TEXT_BOX":
            el_type = "text_box"
        else:
            el_type = f"shape ({shape_type})"
        text_runs = []
        styles_seen = []
        for te in (el["shape"].get("text", {}).get("textElements", [])):
            run = te.get("textRun", {})
            content = run.get("content", "").strip()
            if content:
                text_runs.append(content)
            style = run.get("style", {})
            if style and content:
                s_parts = []
                font = style.get("fontFamily")
                if font:
                    s_parts.append(font)
                fs = style.get("fontSize", {})
                if fs.get("magnitude"):
                    s_parts.append(f"{fs['magnitude']}{fs.get('unit', 'PT').lower()}")
                if style.get("bold"):
                    s_parts.append("bold")
                if style.get("italic"):
                    s_parts.append("italic")
                fg = style.get("foregroundColor", {}).get("opaqueColor", {}).get("rgbColor", {})
                if fg:
                    r_val = int(fg.get("red", 0) * 255)
                    g_val = int(fg.get("green", 0) * 255)
                    b_val = int(fg.get("blue", 0) * 255)
                    hex_color = f"#{r_val:02x}{g_val:02x}{b_val:02x}"
                    if hex_color != "#000000":
                        s_parts.append(hex_color)
                if s_parts:
                    styles_seen.append(" ".join(s_parts))
        text_content = " ".join(text_runs)
        if styles_seen:
            unique_styles = list(dict.fromkeys(styles_seen))
            text_style_desc = " | ".join(unique_styles[:3])
    elif "image" in el:
        el_type = "image"
    elif "table" in el:
        el_type = "table"
    elif "elementGroup" in el:
        el_type = "group"

    summary = {
        "objectId": el["objectId"],
        "type": el_type,
        "x_pt": round(x_pt, 1),
        "y_pt": round(y_pt, 1),
        "width_pt": round(rendered_w, 1),
        "height_pt": round(rendered_h, 1),
    }
    if text_content:
        summary["text"] = text_content[:100]
    if text_style_desc:
        summary["style"] = text_style_desc
    return summary


def _build_slide_description(
    page_json: dict, page_width_emu: int, page_height_emu: int
) -> str:
    """Build a human-readable description of the slide for the LLM."""
    page_w_pt = round(page_width_emu / PT_TO_EMU, 1)
    page_h_pt = round(page_height_emu / PT_TO_EMU, 1)

    elements = page_json.get("pageElements", [])
    summaries = []
    for el in elements:
        s = _summarize_element(el)
        if s:
            summaries.append(s)

    lines = [
        f"Slide dimensions: {page_w_pt} x {page_h_pt} PT (points). Origin (0,0) is top-left.",
        f"Number of elements: {len(summaries)}",
        "",
        "Elements:",
    ]
    for s in summaries:
        desc = f'  - objectId: "{s["objectId"]}", type: {s["type"]}, position: ({s["x_pt"]}, {s["y_pt"]}) PT, size: {s["width_pt"]} x {s["height_pt"]} PT'
        if s.get("text"):
            desc += f', text: "{s["text"]}"'
        if s.get("style"):
            desc += f', style: [{s["style"]}]'
        lines.append(desc)

    return "\n".join(lines)


def _summarize_slide_brief(slide: dict, index: int) -> str:
    """Short summary of a slide (for non-current slides), including text style."""
    elements = slide.get("pageElements", [])
    texts = []
    styles = []
    for el in elements:
        for te in el.get("shape", {}).get("text", {}).get("textElements", []):
            run = te.get("textRun", {})
            t = run.get("content", "").strip()
            if t:
                texts.append(t)
            style = run.get("style", {})
            if style and t:
                s_parts = []
                font = style.get("fontFamily")
                if font:
                    s_parts.append(font)
                fs = style.get("fontSize", {})
                if fs.get("magnitude"):
                    s_parts.append(f"{fs['magnitude']}{fs.get('unit', 'PT').lower()}")
                if style.get("bold"):
                    s_parts.append("bold")
                if s_parts:
                    styles.append(" ".join(s_parts))
    combined = " | ".join(texts)
    if len(combined) > 150:
        combined = combined[:150] + "..."
    el_count = len(elements)
    unique_styles = list(dict.fromkeys(styles))
    style_str = f" (style: {', '.join(unique_styles[:2])})" if unique_styles else ""
    if combined:
        return f"  Slide {index} ({el_count} elements): \"{combined}\"{style_str}"
    return f"  Slide {index} ({el_count} elements): [no text]{style_str}"


def _build_full_presentation_context(
    presentation: dict, current_page_id: Optional[str], page_width_emu: int, page_height_emu: int,
) -> tuple[str, Optional[dict], int, Optional[int]]:
    """
    Build context string with ALL slides summarized.
    Returns (context_string, current_page_json, total_slides, current_slide_index).
    """
    slides = presentation.get("slides", [])
    total = len(slides)
    current_index = None
    current_page_json = None

    # Find current slide
    for i, s in enumerate(slides):
        if s.get("objectId") == current_page_id:
            current_index = i
            current_page_json = s
            break

    # Build other slides summary
    other_lines = []
    for i, s in enumerate(slides):
        if i == current_index:
            continue
        other_lines.append(_summarize_slide_brief(s, i))

    # Build current slide detail
    if current_page_json:
        current_desc = _build_slide_description(current_page_json, page_width_emu, page_height_emu)
    else:
        current_desc = "(Could not find the current slide)"

    lines = [
        f"=== PRESENTATION ({total} slides) ===",
        f"Title: {presentation.get('title', 'Untitled')}",
        "",
        f"=== CURRENT SLIDE (index {current_index}, the slide the user is viewing) ===",
        current_desc,
        "",
        f"=== OTHER SLIDES (summary) ===",
    ]
    lines.extend(other_lines if other_lines else ["  (no other slides)"])

    return "\n".join(lines), current_page_json, total, current_index


SLIDES_LLM_SYSTEM = """You are a Google Slides layout assistant. You receive a detailed description of the current slide's elements, a summary of ALL other slides in the presentation, and a user request. You output JSON instructions.

*** CRITICAL RULE — create_text_box vs create_slide ***
- DEFAULT to "create_text_box" when the user wants to ADD content to the CURRENT slide (e.g. "add a section", "add text about X", "add a box at the bottom", "put a summary here", "add content about Y").
- ONLY use "create_slide" when the user EXPLICITLY says "new slide", "create a slide", "add a slide", or "next slide".
- If ambiguous, ALWAYS prefer create_text_box on the current slide.
***

Supported instruction types:

1. Layout instructions (edit existing elements):
- "action": one of "move", "resize", or "move_and_resize"
- "objectId": the element's objectId (string)
- "x_pt", "y_pt": new position in points (for move/move_and_resize)
- "width_pt", "height_pt": new size in points (for resize/move_and_resize)

2. Create text box (add a NEW text box to the CURRENT slide):
- "action": "create_text_box"
- "x_pt": X position in points
- "y_pt": Y position in points
- "width_pt": width in points
- "height_pt": height in points
- "text": text content (string)
- Optional style: "font_size_pt", "bold", "italic", "underline", "font_family", "color" (hex)

3. Create slide (add a BRAND NEW SLIDE — only when user explicitly asks):
- "action": "create_slide"
- "layout": one of "BLANK", "TITLE", "TITLE_AND_BODY", "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY", "SECTION_HEADER", "CAPTION_ONLY", "BIG_NUMBER"
- "insert_after": "current", "end", or a 0-based index number
- "title": optional title text
- "body": optional body text

4. Replace text (replace ALL text in an existing element):
- "action": "replace_text"
- "objectId": the element's objectId (string)
- "new_text": the replacement text (string)

5. Update text style (change formatting of ALL text in an existing element):
- "action": "update_text_style"
- "objectId": the element's objectId (string)
- Style fields (include one or more): "font_size_pt", "bold", "italic", "underline", "font_family", "color" (hex)

Rules:
- Positions are from the top-left corner of the slide (origin 0,0).
- Only include elements you want to change.
- Be precise with numbers. Think about centering, alignment, and spacing.
- The slide center X is half the slide width. The slide center Y is half the slide height.
- For create_text_box: position it to avoid overlapping existing elements. Match the slide's existing font/style when possible.
- For create_slide: choose a fitting layout. Generate appropriate title/body text for the topic.
- For replace_text: replaces ALL text in the element.
- For update_text_style: applies to ALL text in the element.
- You can combine multiple instructions (e.g. create_text_box + move existing elements to make room).
- When matching formatting across slides, use style info from other slides' summaries.

Output format:
- Output a JSON OBJECT with two keys:
  - "instructions": JSON array of instructions (empty [] if no changes needed)
  - "message": brief natural-language summary of what you did or answered
- No markdown fences. Output ONLY the JSON object.
- For questions about the presentation, output empty instructions and put the answer in "message"."""


def _ask_llm_for_instructions(
    slide_description: str, user_message: str
) -> tuple[list[dict], str]:
    """
    Ask the LLM to generate layout instructions given the slide and user request.
    Returns (instructions_list, message_string).
    """
    messages = [
        SystemMessage(content=SLIDES_LLM_SYSTEM),
        HumanMessage(content=f"Slide structure:\n{slide_description}\n\nUser request: {user_message}"),
    ]

    response = llm.invoke(messages)
    text = response.content if hasattr(response, "content") else str(response)
    text = text.strip()

    # Strip markdown code fences if present
    if text.startswith("```"):
        text = re.sub(r"^```[a-z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()

    print(f"   SLIDES LLM raw response: {text[:500]}")

    try:
        parsed = json.loads(text)
        # New format: {"instructions": [...], "message": "..."}
        if isinstance(parsed, dict) and "instructions" in parsed:
            instructions = parsed.get("instructions", [])
            llm_message = parsed.get("message", "")
            if not isinstance(instructions, list):
                instructions = []
            return instructions, llm_message
        # Fallback: old format (raw array)
        if isinstance(parsed, list):
            return parsed, ""
        print(f"   SLIDES LLM returned unexpected type: {type(parsed)}")
        return [], ""
    except json.JSONDecodeError as e:
        print(f"   SLIDES LLM JSON parse error: {e}")
        return [], ""


def _gen_id(prefix: str = "dex2") -> str:
    """Generate a short unique objectId valid for Slides API (5-50 chars, alphanumeric/underscore start)."""
    return f"{prefix}_{uuid.uuid4().hex[:8]}"


def _resolve_insertion_index(insert_after, current_slide_index: Optional[int], total_slides: int) -> int:
    """Compute the 0-based insertion index for a new slide."""
    if insert_after == "end":
        return total_slides
    elif insert_after == "current" and current_slide_index is not None:
        return current_slide_index + 1
    elif isinstance(insert_after, (int, float)):
        return int(insert_after)
    return total_slides


def _create_and_populate_slide(
    inst: dict,
    presentation_id: str,
    access_token: str,
    current_slide_index: Optional[int],
    total_slides: int,
) -> str:
    """
    Create a slide and populate its placeholders with text.
    Uses two API calls: one to create the slide, one to read it back and insert text.
    Returns a status message.
    """
    layout = inst.get("layout", "BLANK")
    insertion_index = _resolve_insertion_index(
        inst.get("insert_after", "current"), current_slide_index, total_slides
    )
    slide_id = _gen_id("slide")

    create_req = {
        "createSlide": {
            "objectId": slide_id,
            "insertionIndex": insertion_index,
            "slideLayoutReference": {"predefinedLayout": layout},
        }
    }
    execute_batch_update(presentation_id, [create_req], access_token)

    title_text = inst.get("title", "")
    body_text = inst.get("body", "")
    if not title_text and not body_text:
        return f"Created blank slide ({layout})"

    page_json = get_slide_elements(presentation_id, slide_id, access_token)
    text_requests = []
    for el in page_json.get("pageElements", []):
        ph = el.get("shape", {}).get("placeholder", {})
        ph_type = ph.get("type", "")
        obj_id = el.get("objectId")
        if not obj_id:
            continue
        if ph_type in ("TITLE", "CENTERED_TITLE") and title_text:
            text_requests.append({
                "insertText": {"objectId": obj_id, "text": title_text, "insertionIndex": 0}
            })
        elif ph_type in ("BODY", "SUBTITLE") and body_text:
            text_requests.append({
                "insertText": {"objectId": obj_id, "text": body_text, "insertionIndex": 0}
            })

    if text_requests:
        execute_batch_update(presentation_id, text_requests, access_token)

    return f"Created slide ({layout}) with content"


def _create_text_box(
    inst: dict,
    presentation_id: str,
    page_id: str,
    access_token: str,
) -> str:
    """
    Create a text box on the specified slide with optional text and styling.
    Returns a status message.
    """
    obj_id = _gen_id("txtbox")
    x_pt = inst.get("x_pt", 50)
    y_pt = inst.get("y_pt", 50)
    width_pt = inst.get("width_pt", 400)
    height_pt = inst.get("height_pt", 100)
    text = inst.get("text", "")

    requests = [
        {
            "createShape": {
                "objectId": obj_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {
                        "width": {"magnitude": width_pt, "unit": "PT"},
                        "height": {"magnitude": height_pt, "unit": "PT"},
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "shearX": 0,
                        "shearY": 0,
                        "translateX": x_pt * PT_TO_EMU,
                        "translateY": y_pt * PT_TO_EMU,
                        "unit": "EMU",
                    },
                },
            }
        }
    ]

    if text:
        requests.append({
            "insertText": {
                "objectId": obj_id,
                "text": text,
                "insertionIndex": 0,
            }
        })

    style = {}
    fields = []
    if "font_size_pt" in inst:
        style["fontSize"] = {"magnitude": inst["font_size_pt"], "unit": "PT"}
        fields.append("fontSize")
    if "bold" in inst:
        style["bold"] = inst["bold"]
        fields.append("bold")
    if "italic" in inst:
        style["italic"] = inst["italic"]
        fields.append("italic")
    if "underline" in inst:
        style["underline"] = inst["underline"]
        fields.append("underline")
    if "font_family" in inst:
        style["fontFamily"] = inst["font_family"]
        fields.append("fontFamily")
    if "color" in inst:
        style["foregroundColor"] = {
            "opaqueColor": {"rgbColor": _hex_to_rgb(inst["color"])}
        }
        fields.append("foregroundColor")

    if style and fields and text:
        requests.append({
            "updateTextStyle": {
                "objectId": obj_id,
                "textRange": {"type": "ALL"},
                "style": style,
                "fields": ",".join(fields),
            }
        })

    execute_batch_update(presentation_id, requests, access_token)
    return f"Created text box '{text[:40]}...' at ({x_pt}, {y_pt})" if len(text) > 40 else f"Created text box '{text}' at ({x_pt}, {y_pt})"


def _hex_to_rgb(hex_color: str) -> dict:
    """Convert '#RRGGBB' to Slides API rgbColor (0.0-1.0 floats)."""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return {"red": 0, "green": 0, "blue": 0}
    return {
        "red": int(h[0:2], 16) / 255.0,
        "green": int(h[2:4], 16) / 255.0,
        "blue": int(h[4:6], 16) / 255.0,
    }


def _edit_instructions_to_batch_requests(
    instructions: list[dict], page_json: dict,
) -> list[dict]:
    """
    Convert LLM instructions (move/resize/replace_text/update_text_style)
    to Slides API batchUpdate requests.
    Skips create_slide instructions (handled separately).
    """
    el_map = {}
    for el in page_json.get("pageElements", []):
        el_map[el["objectId"]] = el

    requests = []
    for inst in instructions:
        action = inst.get("action")

        if action in ("create_slide", "create_text_box"):
            continue

        obj_id = inst.get("objectId")
        if not obj_id or obj_id not in el_map:
            continue

        # Check if element has text before text operations
        el = el_map[obj_id]
        has_text = bool(
            el.get("shape", {}).get("text", {}).get("textElements")
        )

        if action == "replace_text":
            new_text = inst.get("new_text", "")
            if has_text:
                requests.append({
                    "deleteText": {
                        "objectId": obj_id,
                        "textRange": {"type": "ALL"},
                    }
                })
            requests.append({
                "insertText": {
                    "objectId": obj_id,
                    "text": new_text,
                    "insertionIndex": 0,
                }
            })
            continue

        if action == "update_text_style":
            if not has_text:
                print(f"   SLIDES: skipping update_text_style for {obj_id} (no text)")
                continue
            style = {}
            fields = []

            if "font_size_pt" in inst:
                style["fontSize"] = {"magnitude": inst["font_size_pt"], "unit": "PT"}
                fields.append("fontSize")
            if "bold" in inst:
                style["bold"] = inst["bold"]
                fields.append("bold")
            if "italic" in inst:
                style["italic"] = inst["italic"]
                fields.append("italic")
            if "underline" in inst:
                style["underline"] = inst["underline"]
                fields.append("underline")
            if "font_family" in inst:
                style["fontFamily"] = inst["font_family"]
                fields.append("fontFamily")
            if "color" in inst:
                style["foregroundColor"] = {
                    "opaqueColor": {"rgbColor": _hex_to_rgb(inst["color"])}
                }
                fields.append("foregroundColor")

            if style and fields:
                requests.append({
                    "updateTextStyle": {
                        "objectId": obj_id,
                        "textRange": {"type": "ALL"},
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })
            continue

        # Layout instructions (move / resize / move_and_resize)
        transform = el.get("transform", {})
        size = el.get("size", {})

        existing_sx = transform.get("scaleX", 1)
        existing_sy = transform.get("scaleY", 1)
        existing_shx = transform.get("shearX", 0)
        existing_shy = transform.get("shearY", 0)
        existing_tx = transform.get("translateX", 0)
        existing_ty = transform.get("translateY", 0)

        new_tx = existing_tx
        new_ty = existing_ty
        new_sx = existing_sx
        new_sy = existing_sy

        if action in ("move", "move_and_resize"):
            if "x_pt" in inst:
                new_tx = inst["x_pt"] * PT_TO_EMU
            if "y_pt" in inst:
                new_ty = inst["y_pt"] * PT_TO_EMU

        if action in ("resize", "move_and_resize"):
            w_dim = size.get("width", {})
            h_dim = size.get("height", {})
            raw_w = w_dim.get("magnitude", 1)
            raw_h = h_dim.get("magnitude", 1)
            if w_dim.get("unit") == "PT":
                raw_w_pt = raw_w
            else:
                raw_w_pt = raw_w / PT_TO_EMU
            if h_dim.get("unit") == "PT":
                raw_h_pt = raw_h
            else:
                raw_h_pt = raw_h / PT_TO_EMU

            if "width_pt" in inst and raw_w_pt > 0:
                new_sx = inst["width_pt"] / raw_w_pt
            if "height_pt" in inst and raw_h_pt > 0:
                new_sy = inst["height_pt"] / raw_h_pt

        requests.append({
            "updatePageElementTransform": {
                "objectId": obj_id,
                "applyMode": "ABSOLUTE",
                "transform": {
                    "scaleX": new_sx,
                    "scaleY": new_sy,
                    "shearX": existing_shx,
                    "shearY": existing_shy,
                    "translateX": new_tx,
                    "translateY": new_ty,
                    "unit": "EMU",
                },
            }
        })

    return requests


def execute_batch_update(
    presentation_id: str, requests: list[dict], access_token: str
) -> dict:
    """Send a batchUpdate to the Slides API."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    r = httpx.post(
        f"{SLIDES_API}/{presentation_id}:batchUpdate",
        headers=headers,
        json={"requests": requests},
        timeout=30.0,
    )
    r.raise_for_status()
    return r.json()


def handle_edit_slides(
    current_tab_url: Optional[str],
    user_message: str,
    access_token: Optional[str],
) -> str:
    """
    Orchestrator: parse URL, read slide, ask LLM for layout instructions,
    translate to API calls, execute.
    """
    if not access_token:
        return "I need Google access to edit your slides. Please click Connect Google in the extension and try again."

    if not current_tab_url:
        return "I can't detect which presentation you're viewing. Please open a Google Slides tab and try again."

    parsed = parse_slides_url(current_tab_url)
    if not parsed:
        return "The current tab doesn't appear to be a Google Slides presentation. Please open a slide and try again."

    presentation_id, page_id = parsed

    if not page_id:
        return "I can't tell which slide you're on. Click on a slide so the URL shows #slide=id.XXX, then try again."

    try:
        presentation = get_full_presentation(presentation_id, access_token)
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 401:
            return "Google token expired or missing Slides permission. Please Disconnect Google, then Connect Google again."
        return f"Failed to read the presentation (HTTP {e.response.status_code}). Check that the Slides API is enabled in Google Cloud."
    except Exception as e:
        return f"Error reading presentation: {e}"

    page_size = presentation.get("pageSize", {})
    page_width = int(page_size.get("width", {}).get("magnitude", DEFAULT_PAGE_WIDTH_EMU))
    page_height = int(page_size.get("height", {}).get("magnitude", DEFAULT_PAGE_HEIGHT_EMU))

    full_desc, page_json, total_slides, current_slide_index = _build_full_presentation_context(
        presentation, page_id, page_width, page_height,
    )

    if not page_json:
        return "I can't find the current slide in the presentation. Try clicking on the slide and sending your request again."

    print(f"   SLIDES: Full presentation context:\n{full_desc[:2000]}{'...(truncated)' if len(full_desc) > 2000 else ''}")

    try:
        instructions, llm_message = _ask_llm_for_instructions(full_desc, user_message)
    except Exception as e:
        return f"Error getting layout instructions from LLM: {e}"

    if not instructions:
        if llm_message:
            return llm_message
        return "The AI couldn't determine any changes for this request. Try being more specific (e.g. 'create a title slide about AI' or 'center the two text boxes')."

    print(f"   SLIDES: {len(instructions)} instructions from LLM")

    slides_created = 0
    elements_updated = 0
    textboxes_created = 0

    # Handle create_text_box instructions
    textbox_instructions = [i for i in instructions if i.get("action") == "create_text_box"]
    for ti in textbox_instructions:
        try:
            result = _create_text_box(ti, presentation_id, page_id, access_token)
            print(f"   SLIDES: {result}")
            textboxes_created += 1
        except httpx.HTTPStatusError as e:
            error_body = ""
            try:
                error_body = e.response.text[:300]
            except Exception:
                pass
            print(f"   SLIDES: create_text_box failed (HTTP {e.response.status_code}): {error_body}")
            return f"Failed to create text box (HTTP {e.response.status_code}). Error: {error_body}"
        except Exception as e:
            return f"Error creating text box: {e}"

    # Handle create_slide instructions
    create_instructions = [i for i in instructions if i.get("action") == "create_slide"]
    for ci in create_instructions:
        try:
            result = _create_and_populate_slide(
                ci, presentation_id, access_token, current_slide_index, total_slides,
            )
            print(f"   SLIDES: {result}")
            slides_created += 1
        except httpx.HTTPStatusError as e:
            error_body = ""
            try:
                error_body = e.response.text[:300]
            except Exception:
                pass
            print(f"   SLIDES: create_slide failed (HTTP {e.response.status_code}): {error_body}")
            return f"Failed to create the slide (HTTP {e.response.status_code}). Error: {error_body}"
        except Exception as e:
            return f"Error creating slide: {e}"

    # Handle edit instructions (move/resize/replace_text/update_text_style)
    layout_requests = _edit_instructions_to_batch_requests(instructions, page_json)
    if layout_requests:
        try:
            execute_batch_update(presentation_id, layout_requests, access_token)
            elements_updated = len(layout_requests)
        except httpx.HTTPStatusError as e:
            error_body = ""
            try:
                error_body = e.response.text[:300]
            except Exception:
                pass
            print(f"   SLIDES: edit requests failed (HTTP {e.response.status_code}): {error_body}")
            return f"Failed to update elements (HTTP {e.response.status_code}). Make sure the Slides API is enabled and you have edit access."
        except Exception as e:
            return f"Error updating elements: {e}"

    if not slides_created and not elements_updated and not textboxes_created:
        if llm_message:
            return llm_message
        return "No valid changes could be generated. Try being more specific."

    parts = []
    if textboxes_created:
        parts.append(f"added {textboxes_created} text box{'es' if textboxes_created != 1 else ''}")
    if slides_created:
        parts.append(f"created {slides_created} new slide{'s' if slides_created != 1 else ''}")
    if elements_updated:
        parts.append(f"updated {elements_updated} element{'s' if elements_updated != 1 else ''}")
    summary = " and ".join(parts)
    result_msg = f"Done! I {summary}. Refresh your Slides tab to see the changes."
    if llm_message:
        result_msg += f"\n\n{llm_message}"
    return result_msg
