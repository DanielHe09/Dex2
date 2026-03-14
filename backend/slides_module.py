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


def get_slide_list(presentation_id: str, access_token: str) -> list[dict]:
    """Return list of slide objectIds in order. Each entry: {"objectId": "...", "index": 0}."""
    headers = {"Authorization": f"Bearer {access_token}"}
    r = httpx.get(
        f"{SLIDES_API}/{presentation_id}",
        headers=headers,
        params={"fields": "slides.objectId"},
        timeout=15.0,
    )
    r.raise_for_status()
    slides = r.json().get("slides", [])
    return [{"objectId": s["objectId"], "index": i} for i, s in enumerate(slides)]


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

    # Determine element type and text content
    el_type = "shape"
    text_content = ""
    if "shape" in el:
        shape_type = el["shape"].get("shapeType", "RECTANGLE")
        if shape_type == "TEXT_BOX":
            el_type = "text_box"
        else:
            el_type = f"shape ({shape_type})"
        text_runs = []
        for te in (el["shape"].get("text", {}).get("textElements", [])):
            run = te.get("textRun", {})
            if run.get("content", "").strip():
                text_runs.append(run["content"].strip())
        text_content = " ".join(text_runs)
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
        lines.append(desc)

    return "\n".join(lines)


SLIDES_LLM_SYSTEM = """You are a Google Slides layout assistant. You receive a description of a slide's elements (positions, sizes, types), presentation context (total slides, current slide index), and a user request. You output a JSON array of instructions.

Supported instruction types:

1. Layout instructions (edit existing elements):
- "action": one of "move", "resize", or "move_and_resize"
- "objectId": the element's objectId (string)
- "x_pt": new X position in points (for move/move_and_resize)
- "y_pt": new Y position in points (for move/move_and_resize)
- "width_pt": new width in points (for resize/move_and_resize)
- "height_pt": new height in points (for resize/move_and_resize)

2. Create slide instruction:
- "action": "create_slide"
- "layout": one of "BLANK", "TITLE", "TITLE_AND_BODY", "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY", "SECTION_HEADER", "CAPTION_ONLY", "BIG_NUMBER"
- "insert_after": "current" (after the slide the user is viewing) or "end" (at the end) or a 0-based index number
- "title": optional title text for the slide
- "body": optional body text for the slide

3. Replace text instruction (replace ALL text in an element):
- "action": "replace_text"
- "objectId": the element's objectId (string)
- "new_text": the replacement text (string)

4. Update text style instruction (change formatting of ALL text in an element):
- "action": "update_text_style"
- "objectId": the element's objectId (string)
- Include one or more of these optional style fields:
  - "font_size_pt": font size in points (number, e.g. 24)
  - "bold": true or false
  - "italic": true or false
  - "underline": true or false
  - "font_family": font name (string, e.g. "Arial", "Times New Roman", "Roboto")
  - "color": hex color string (e.g. "#FF0000" for red, "#0000FF" for blue, "#FFFFFF" for white)

Rules:
- Positions are from the top-left corner of the slide (origin 0,0).
- Only include elements you want to change. Do NOT include elements that should stay where they are.
- Output ONLY the JSON array, no explanation, no markdown fences.
- If no changes are needed, output an empty array: []
- Be precise with numbers. Think about centering, alignment, and spacing.
- When making elements symmetrical, consider both their positions AND sizes relative to the slide center.
- The slide center X is half the slide width. The slide center Y is half the slide height.
- For create_slide: choose a layout that fits the user's request. If they ask for a title slide, use "TITLE". If they want a slide with content, use "TITLE_AND_BODY". If they don't specify, use "BLANK".
- You can combine create_slide with layout instructions in the same array (e.g. create a slide then move elements on the current slide).
- When the user says "create a slide about X", generate appropriate title and body text for the topic.
- For replace_text: this replaces ALL text in the element. Use it when the user wants to change what text says.
- For update_text_style: this applies to ALL text in the element. Use it for font size, color, bold, italic, underline, font family changes.
- You can combine update_text_style with replace_text on the same element (e.g. change text AND make it bold).
- When the user says "make all text bigger" or "change font size", apply update_text_style to every text element on the slide."""


def _ask_llm_for_instructions(
    slide_description: str, user_message: str
) -> list[dict]:
    """
    Ask the LLM to generate layout instructions given the slide and user request.
    Returns a list of instruction dicts.
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
        instructions = json.loads(text)
        if not isinstance(instructions, list):
            print(f"   SLIDES LLM returned non-list: {type(instructions)}")
            return []
        return instructions
    except json.JSONDecodeError as e:
        print(f"   SLIDES LLM JSON parse error: {e}")
        return []


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

        if action == "create_slide":
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
        page_width, page_height = get_page_size(presentation_id, access_token)
        page_json = get_slide_elements(presentation_id, page_id, access_token)
        slide_list = get_slide_list(presentation_id, access_token)
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 401:
            return "Google token expired or missing Slides permission. Please Disconnect Google, then Connect Google again."
        return f"Failed to read the slide (HTTP {e.response.status_code}). Check that the Slides API is enabled in Google Cloud."
    except Exception as e:
        return f"Error reading slide: {e}"

    current_slide_index = None
    for s in slide_list:
        if s["objectId"] == page_id:
            current_slide_index = s["index"]
            break
    total_slides = len(slide_list)

    slide_desc = _build_slide_description(page_json, page_width, page_height)
    pres_context = f"\nPresentation has {total_slides} slide(s). Current slide is index {current_slide_index} (0-based)."
    full_desc = slide_desc + pres_context
    print(f"   SLIDES: Slide description:\n{full_desc}")

    try:
        instructions = _ask_llm_for_instructions(full_desc, user_message)
    except Exception as e:
        return f"Error getting layout instructions from LLM: {e}"

    if not instructions:
        return "The AI couldn't determine any changes for this request. Try being more specific (e.g. 'create a title slide about AI' or 'center the two text boxes')."

    print(f"   SLIDES: {len(instructions)} instructions from LLM")

    slides_created = 0
    elements_updated = 0

    # Handle create_slide instructions first
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

    if not slides_created and not elements_updated:
        return "No valid changes could be generated. Try being more specific."

    parts = []
    if slides_created:
        parts.append(f"created {slides_created} new slide{'s' if slides_created != 1 else ''}")
    if elements_updated:
        parts.append(f"updated {elements_updated} element{'s' if elements_updated != 1 else ''}")
    summary = " and ".join(parts)
    return f"Done! I {summary}. Refresh your Slides tab to see the changes."
