"""
Google Slides API module: read slide structure, compute layout changes, apply updates.
"""

import re
from typing import Optional
import httpx


SLIDES_API = "https://slides.googleapis.com/v1/presentations"

# Default slide dimensions in EMU (10 x 7.5 inches)
DEFAULT_PAGE_WIDTH_EMU = 9_144_000
DEFAULT_PAGE_HEIGHT_EMU = 6_858_000


def parse_slides_url(url: str) -> Optional[tuple[str, Optional[str]]]:
    """
    Extract (presentation_id, page_object_id) from a Google Slides URL.
    page_object_id may be None if no specific slide is selected in the URL.
    """
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
    """Return (width_emu, height_emu) for the presentation, falling back to defaults."""
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
    """
    Fetch a single slide's page elements via presentations.pages.get.
    Returns the raw page JSON from the Slides API.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    r = httpx.get(
        f"{SLIDES_API}/{presentation_id}/pages/{page_id}",
        headers=headers,
        timeout=15.0,
    )
    r.raise_for_status()
    return r.json()


def _element_bounds(el: dict) -> Optional[dict]:
    """
    Extract bounding info for a page element.
    Returns dict with objectId, translateX, translateY, width, height (all in EMU)
    or None if transform is missing.
    """
    transform = el.get("transform")
    size = el.get("size", {})
    if not transform:
        return None
    w_dim = size.get("width", {})
    h_dim = size.get("height", {})
    w = w_dim.get("magnitude", 0)
    h = h_dim.get("magnitude", 0)
    # EMU conversion: if unit is PT, convert (1 PT = 12700 EMU)
    if w_dim.get("unit") == "PT":
        w *= 12700
    if h_dim.get("unit") == "PT":
        h *= 12700
    scale_x = transform.get("scaleX", 1)
    scale_y = transform.get("scaleY", 1)
    return {
        "objectId": el["objectId"],
        "translateX": transform.get("translateX", 0),
        "translateY": transform.get("translateY", 0),
        "scaleX": scale_x,
        "scaleY": scale_y,
        "shearX": transform.get("shearX", 0),
        "shearY": transform.get("shearY", 0),
        "width": w * abs(scale_x),
        "height": h * abs(scale_y),
        "raw_width": w,
        "raw_height": h,
    }


def make_symmetrical(page_json: dict, page_width: int, page_height: int) -> list[dict]:
    """
    Deterministic horizontal symmetry: center all movable elements as a group
    around the slide's horizontal midpoint, preserving their relative spacing
    and vertical positions.

    Returns a list of UpdatePageElementTransformRequest dicts for batchUpdate.
    """
    elements = page_json.get("pageElements", [])
    bounds = []
    for el in elements:
        b = _element_bounds(el)
        if b:
            bounds.append(b)

    if not bounds:
        return []

    center_x = page_width / 2

    group_left = min(b["translateX"] for b in bounds)
    group_right = max(b["translateX"] + b["width"] for b in bounds)
    group_width = group_right - group_left
    group_center = group_left + group_width / 2

    shift_x = center_x - group_center

    requests = []
    for b in bounds:
        new_translate_x = b["translateX"] + shift_x
        requests.append({
            "updatePageElementTransform": {
                "objectId": b["objectId"],
                "applyMode": "ABSOLUTE",
                "transform": {
                    "scaleX": b["scaleX"],
                    "scaleY": b["scaleY"],
                    "shearX": b["shearX"],
                    "shearY": b["shearY"],
                    "translateX": new_translate_x,
                    "translateY": b["translateY"],
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
    Orchestrator: parse the Slides URL, read the slide, compute changes, apply them.
    Returns a user-facing message string.
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

    update_requests = make_symmetrical(page_json, page_width, page_height)

    if not update_requests:
        return "No movable elements found on this slide."

    try:
        execute_batch_update(presentation_id, update_requests, access_token)
    except httpx.HTTPStatusError as e:
        return f"Failed to update the slide (HTTP {e.response.status_code}). Make sure the Slides API is enabled and you have edit access."
    except Exception as e:
        return f"Error updating slide: {e}"

    count = len(update_requests)
    return f"Done! I centered {count} element{'s' if count != 1 else ''} symmetrically on the slide. Refresh your Slides tab to see the changes."
