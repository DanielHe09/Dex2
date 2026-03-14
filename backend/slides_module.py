"""
Google Slides API module: read slide structure, use LLM to plan layout changes,
translate to Slides API batchUpdate requests.
"""

import json
import re
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


SLIDES_LLM_SYSTEM = """You are a Google Slides layout assistant. You receive a description of a slide's elements (positions, sizes, types) and a user request. You output a JSON array of layout instructions.

Each instruction is an object with these fields:
- "action": one of "move", "resize", or "move_and_resize"
- "objectId": the element's objectId (string)
- "x_pt": new X position in points (for move/move_and_resize)
- "y_pt": new Y position in points (for move/move_and_resize)
- "width_pt": new width in points (for resize/move_and_resize)
- "height_pt": new height in points (for resize/move_and_resize)

Rules:
- Positions are from the top-left corner of the slide (origin 0,0).
- Only include elements you want to change. Do NOT include elements that should stay where they are.
- Output ONLY the JSON array, no explanation, no markdown fences.
- If no changes are needed, output an empty array: []
- Be precise with numbers. Think about centering, alignment, and spacing.
- When making elements symmetrical, consider both their positions AND sizes relative to the slide center.
- The slide center X is half the slide width. The slide center Y is half the slide height."""


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


def _instructions_to_batch_requests(
    instructions: list[dict], page_json: dict
) -> list[dict]:
    """
    Convert LLM layout instructions to Slides API batchUpdate requests.
    Preserves existing transform values (scale, shear) for elements.
    """
    el_map = {}
    for el in page_json.get("pageElements", []):
        el_map[el["objectId"]] = el

    requests = []
    for inst in instructions:
        action = inst.get("action")
        obj_id = inst.get("objectId")
        if not obj_id or obj_id not in el_map:
            continue

        el = el_map[obj_id]
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
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 401:
            return "Google token expired or missing Slides permission. Please Disconnect Google, then Connect Google again."
        return f"Failed to read the slide (HTTP {e.response.status_code}). Check that the Slides API is enabled in Google Cloud."
    except Exception as e:
        return f"Error reading slide: {e}"

    slide_desc = _build_slide_description(page_json, page_width, page_height)
    print(f"   SLIDES: Slide description:\n{slide_desc}")

    try:
        instructions = _ask_llm_for_instructions(slide_desc, user_message)
    except Exception as e:
        return f"Error getting layout instructions from LLM: {e}"

    if not instructions:
        return "The AI couldn't determine any layout changes for this request. Try being more specific (e.g. 'center the two text boxes horizontally')."

    print(f"   SLIDES: {len(instructions)} instructions from LLM")

    update_requests = _instructions_to_batch_requests(instructions, page_json)

    if not update_requests:
        return "No valid layout changes could be generated. The element IDs from the AI might not match the slide."

    try:
        execute_batch_update(presentation_id, update_requests, access_token)
    except httpx.HTTPStatusError as e:
        return f"Failed to update the slide (HTTP {e.response.status_code}). Make sure the Slides API is enabled and you have edit access."
    except Exception as e:
        return f"Error updating slide: {e}"

    count = len(update_requests)
    return f"Done! I updated {count} element{'s' if count != 1 else ''} on the slide. Refresh your Slides tab to see the changes."
