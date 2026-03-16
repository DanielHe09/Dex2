"""
Main entry point: two-step router → executor orchestration for Google Slides editing.
"""

from typing import Optional

import httpx

from .api import (
    DEFAULT_PAGE_HEIGHT_EMU, DEFAULT_PAGE_WIDTH_EMU, PT_TO_EMU,
    execute_batch_update, gen_id, get_full_presentation, parse_slides_url,
    resolve_insertion_index,
)
from .context import (
    build_full_presentation_context, extract_presentation_style,
    get_presentation_style_values,
)
from . import vision_style
from .router import build_router_context, route_request
from .executors import (
    ANSWER_QUESTION_PROMPT, CREATE_CONTENT_PROMPT, CREATE_SLIDE_PROMPT,
    EDIT_LAYOUT_PROMPT, EDIT_TEXT_PROMPT, call_executor,
)
from .actions import apply_instructions
from .layout import prepare_instructions_for_apply, normalize_instructions_style


def handle_edit_slides(
    current_tab_url: Optional[str],
    user_message: str,
    access_token: Optional[str],
    slide_screenshot: Optional[str] = None,
) -> str:
    """
    Two-step orchestrator:
      1. ROUTER — classify user request into an operation type
      2. EXECUTOR — run the focused prompt for that operation
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

    # Fetch presentation
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
    page_w_pt = round(page_width / PT_TO_EMU, 1)
    page_h_pt = round(page_height / PT_TO_EMU, 1)

    full_desc, page_json, total_slides, current_slide_index, gaps, num_elements = (
        build_full_presentation_context(presentation, page_id, page_width, page_height)
    )
    if not page_json:
        return "I can't find the current slide in the presentation. Try clicking on the slide and sending your request again."

    # --- STEP 1: ROUTE ---
    router_ctx = build_router_context(
        total_slides, presentation.get("title", "Untitled"),
        current_slide_index, page_w_pt, page_h_pt, num_elements, gaps,
    )
    try:
        operation, route_msg = route_request(router_ctx, user_message)
    except Exception as e:
        print(f"   ROUTER error: {e}, falling back to edit_layout")
        operation = "edit_layout"
        route_msg = ""

    print(f"   ROUTED to: {operation} ({route_msg})")

    # --- STEP 2: EXECUTE ---

    if operation == "answer_question":
        _, llm_message, _ = call_executor(ANSWER_QUESTION_PROMPT, full_desc, user_message)
        return llm_message or "I couldn't find an answer. Try rephrasing your question."

    # When adding new elements (create_content) with screenshot, Gemini can generate instructions from the image (placement + style). Otherwise we use GPT executor.
    precomputed_style_values: Optional[dict] = None
    content_from_gemini = False  # if True, normalize only fills missing style fields so Gemini's style is preserved
    if operation == "edit_layout":
        instructions, llm_message, _ = call_executor(EDIT_LAYOUT_PROMPT, full_desc, user_message)

    elif operation == "create_content":
        if slide_screenshot:
            # Gemini analyzes the screenshot and produces instructions (placement + style). Pass full_desc so Gemini knows FREE SPACE gaps and element positions.
            instructions, llm_message = vision_style.generate_content_instructions_from_image(
                slide_screenshot, user_message, page_w_pt, page_h_pt, layout_context=full_desc
            )
            precomputed_style_values = vision_style.extract_style_from_slide_image(slide_screenshot)
            if instructions:
                content_from_gemini = True
            elif precomputed_style_values:
                # Fallback to GPT executor if Gemini returned nothing
                style_blurb = vision_style.format_style_for_prompt(precomputed_style_values)
                create_ctx = f"{full_desc}\n\n{style_blurb}"
                instructions, llm_message, _ = call_executor(CREATE_CONTENT_PROMPT, create_ctx, user_message)
            else:
                instructions, llm_message, _ = call_executor(CREATE_CONTENT_PROMPT, full_desc, user_message)
        else:
            instructions, llm_message, _ = call_executor(CREATE_CONTENT_PROMPT, full_desc, user_message)

    elif operation == "create_slide":
        return _handle_create_slide(
            presentation, full_desc, user_message,
            presentation_id, access_token,
            page_w_pt, page_h_pt, current_slide_index, total_slides,
            slide_screenshot=slide_screenshot,
        )

    elif operation == "edit_text":
        edit_ctx = full_desc + "\n\n" + extract_presentation_style(presentation)
        instructions, llm_message, _ = call_executor(EDIT_TEXT_PROMPT, edit_ctx, user_message)

    else:
        instructions, llm_message, _ = call_executor(EDIT_LAYOUT_PROMPT, full_desc, user_message)

    # Apply instructions for non-create_slide operations
    if not instructions:
        if llm_message:
            return llm_message
        return "The AI couldn't determine any changes for this request. Try being more specific."

    print(f"   SLIDES: {len(instructions)} instructions from executor")

    instructions = prepare_instructions_for_apply(instructions, page_w_pt, page_h_pt)
    # Use vision style when create_content had screenshot; otherwise API-based style. If Gemini generated content, only fill missing style fields.
    if precomputed_style_values is not None:
        style_values = precomputed_style_values
    else:
        style_values = get_presentation_style_values(presentation, presentation_id, access_token)
    instructions = normalize_instructions_style(
        instructions, style_values, fill_missing_only=content_from_gemini
    )

    sc, eu, err = apply_instructions(
        instructions, presentation_id, page_id, page_json, access_token,
    )
    if err:
        return err

    if not sc and not eu:
        if llm_message:
            return llm_message
        return "No valid changes could be generated. Try being more specific."

    parts = []
    if sc:
        parts.append(f"added {sc} element{'s' if sc != 1 else ''}")
    if eu:
        parts.append(f"updated {eu} element{'s' if eu != 1 else ''}")
    summary = " and ".join(parts)
    result_msg = f"Done! I {summary}. Refresh your Slides tab to see the changes."
    if llm_message:
        result_msg += f"\n\n{llm_message}"
    return result_msg


def _handle_create_slide(
    presentation: dict, full_desc: str, user_message: str,
    presentation_id: str, access_token: str,
    page_w_pt: float, page_h_pt: float,
    current_slide_index: Optional[int], total_slides: int,
    slide_screenshot: Optional[str] = None,
) -> str:
    """Handle the create_slide operation: create BLANK slide, populate with styled shapes."""
    style_info = extract_presentation_style(presentation)
    slide_context = f"{full_desc}\n\n{style_info}\n\nNew slide dimensions: {page_w_pt} x {page_h_pt} PT"
    if slide_screenshot:
        vision_style_values = vision_style.extract_style_from_slide_image(slide_screenshot)
        if vision_style_values:
            slide_context = f"{slide_context}\n\n{vision_style.format_style_for_prompt(vision_style_values)}"
    instructions, llm_message, raw_parsed = call_executor(
        CREATE_SLIDE_PROMPT, slide_context, user_message
    )

    if not instructions:
        if llm_message:
            return llm_message
        return "Couldn't generate slide content. Try being more specific."

    insert_after = raw_parsed.get("insert_after", "current")
    insertion_index = resolve_insertion_index(insert_after, current_slide_index, total_slides)
    slide_id = gen_id("slide")

    try:
        execute_batch_update(presentation_id, [{
            "createSlide": {
                "objectId": slide_id,
                "insertionIndex": insertion_index,
                "slideLayoutReference": {"predefinedLayout": "BLANK"},
            }
        }], access_token)
    except httpx.HTTPStatusError as e:
        body = e.response.text[:300] if e.response else ""
        return f"Failed to create slide (HTTP {e.response.status_code}). Error: {body}"

    instructions = prepare_instructions_for_apply(instructions, page_w_pt, page_h_pt)
    if slide_screenshot:
        style_values = vision_style.extract_style_from_slide_image(slide_screenshot) or {}
        if not style_values:
            style_values = get_presentation_style_values(presentation, presentation_id, access_token)
    else:
        style_values = get_presentation_style_values(presentation, presentation_id, access_token)
    instructions = normalize_instructions_style(instructions, style_values)
    empty_page = {"pageElements": []}
    sc, _, err = apply_instructions(
        instructions, presentation_id, slide_id, empty_page, access_token,
    )
    if err:
        return err

    result_msg = f"Done! I created a new slide with {sc} element{'s' if sc != 1 else ''}."
    if llm_message:
        result_msg += f"\n\n{llm_message}"
    return result_msg
