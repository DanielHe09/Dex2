"""
Build human-readable slide/presentation descriptions for LLM context.
"""

from typing import Optional, Any

from .api import PT_TO_EMU, get_slide_elements


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

    width_pt = (w_mag / PT_TO_EMU) if w_unit == "EMU" else w_mag
    height_pt = (h_mag / PT_TO_EMU) if h_unit == "EMU" else h_mag

    scale_x = transform.get("scaleX", 1)
    scale_y = transform.get("scaleY", 1)
    tx = transform.get("translateX", 0)
    ty = transform.get("translateY", 0)

    rendered_w = width_pt * abs(scale_x)
    rendered_h = height_pt * abs(scale_y)
    x_pt = tx / PT_TO_EMU
    y_pt = ty / PT_TO_EMU

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
        for te in el["shape"].get("text", {}).get("textElements", []):
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
        bg_fill = el["shape"].get("shapeProperties", {}).get("shapeBackgroundFill", {})
        solid = bg_fill.get("solidFill", {})
        bg_rgb = solid.get("color", {}).get("rgbColor", {})
        if bg_rgb:
            br = int(bg_rgb.get("red", 1) * 255)
            bg = int(bg_rgb.get("green", 1) * 255)
            bb = int(bg_rgb.get("blue", 1) * 255)
            bg_hex = f"#{br:02x}{bg:02x}{bb:02x}"
            text_style_desc += f" | bg:{bg_hex}"
        outline = el["shape"].get("shapeProperties", {}).get("outline", {})
        if outline:
            out_fill = outline.get("outlineFill", {}).get("solidFill", {}).get("color", {}).get("rgbColor", {})
            if out_fill:
                or_ = int(out_fill.get("red", 0) * 255)
                og = int(out_fill.get("green", 0) * 255)
                ob = int(out_fill.get("blue", 0) * 255)
                out_hex = f"#{or_:02x}{og:02x}{ob:02x}"
                text_style_desc += f" | border:{out_hex}"
            w = outline.get("weight", {}).get("magnitude")
            if w is not None:
                wu = outline.get("weight", {}).get("unit", "PT")
                text_style_desc += f" | border_weight:{w}{wu.lower()}"
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


def list_empty_text_box_summaries(page_json: dict) -> list[dict]:
    """
    Text-box shapes with no visible text (whitespace-only counts as empty).
    Used to route edit_text / replace_text instead of stacking duplicate create_shape boxes.
    """
    summaries: list[dict] = []
    for el in page_json.get("pageElements", []):
        s = _summarize_element(el)
        if not s:
            continue
        if s.get("type") != "text_box":
            continue
        txt = (s.get("text") or "").strip()
        if txt:
            continue
        summaries.append(s)
    return summaries


def _compute_free_gaps(
    summaries: list[dict], page_h_pt: float
) -> list[tuple[float, float, float]]:
    """Compute vertical free-space gaps from element summaries."""
    occupied = []
    for s in summaries:
        if s["width_pt"] > 0 and s["height_pt"] > 0:
            occupied.append((s["y_pt"], s["y_pt"] + s["height_pt"]))
    occupied.sort()
    merged = []
    for start, end in occupied:
        if merged and start <= merged[-1][1] + 2:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    gaps = []
    prev_end = 0.0
    for start, end in merged:
        if start - prev_end >= 15:
            gaps.append((round(prev_end, 1), round(start, 1), round(start - prev_end, 1)))
        prev_end = end
    if page_h_pt - prev_end >= 15:
        gaps.append((round(prev_end, 1), round(page_h_pt, 1), round(page_h_pt - prev_end, 1)))
    return gaps


def build_slide_description(
    page_json: dict, page_width_emu: int, page_height_emu: int
) -> tuple[str, list[tuple[float, float, float]], int]:
    """Build a human-readable description of a slide. Returns (description, gaps, num_elements)."""
    page_w_pt = round(page_width_emu / PT_TO_EMU, 1)
    page_h_pt = round(page_height_emu / PT_TO_EMU, 1)

    elements = page_json.get("pageElements", [])
    summaries = []
    for el in elements:
        s = _summarize_element(el)
        if s:
            summaries.append(s)

    gaps = _compute_free_gaps(summaries, page_h_pt)

    lines = [
        f"Slide dimensions: {page_w_pt} x {page_h_pt} PT (points). Origin (0,0) is top-left.",
        f"Number of elements: {len(summaries)}",
    ]

    if gaps:
        lines.append("")
        lines.append("*** FREE SPACE — place new elements ONLY in these vertical gaps: ***")
        for g_start, g_end, g_height in gaps:
            lines.append(f"  GAP: y={g_start} to y={g_end} (max height: {g_height} PT)")
        lines.append("  A new element at y=Y with height=H MUST satisfy: Y >= gap_start AND Y + H <= gap_end.")
        lines.append("  If your content doesn't fit, SHRINK it (reduce height or font size) to fit within the gap.")

    lines.append("")
    lines.append("Elements:")
    for s in summaries:
        desc = f'  - objectId: "{s["objectId"]}", type: {s["type"]}, position: ({s["x_pt"]}, {s["y_pt"]}) PT, size: {s["width_pt"]} x {s["height_pt"]} PT'
        if s.get("text"):
            desc += f', text: "{s["text"]}"'
        if s.get("style"):
            desc += f', style: [{s["style"]}]'
        lines.append(desc)

    empty_boxes = [s for s in summaries if s.get("type") == "text_box" and not (s.get("text") or "").strip()]
    if empty_boxes:
        lines.append("")
        lines.append(
            "*** EMPTY TEXT BOXES — no visible text. Prefer {\"action\": \"replace_text\", \"objectId\": \"...\", \"new_text\": \"...\"} "
            "to fill one of these; do NOT create_shape a new TEXT_BOX on top (causes stacked duplicates). ***"
        )
        for s in empty_boxes:
            lines.append(
                f'  - objectId: "{s["objectId"]}", position: ({s["x_pt"]}, {s["y_pt"]}) PT, '
                f'size: {s["width_pt"]} x {s["height_pt"]} PT'
            )

    return "\n".join(lines), gaps, len(summaries)


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


def build_full_presentation_context(
    presentation: dict, current_page_id: Optional[str],
    page_width_emu: int, page_height_emu: int,
) -> tuple[str, Optional[dict], int, Optional[int], list, int]:
    """
    Build context string with ALL slides summarized.
    Returns (context_string, current_page_json, total_slides, current_slide_index, gaps, num_elements).
    """
    slides = presentation.get("slides", [])
    total = len(slides)
    current_index = None
    current_page_json = None

    for i, s in enumerate(slides):
        if s.get("objectId") == current_page_id:
            current_index = i
            current_page_json = s
            break

    other_lines = []
    for i, s in enumerate(slides):
        if i == current_index:
            continue
        other_lines.append(_summarize_slide_brief(s, i))

    gaps = []
    num_elements = 0
    if current_page_json:
        current_desc, gaps, num_elements = build_slide_description(
            current_page_json, page_width_emu, page_height_emu
        )
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

    return "\n".join(lines), current_page_json, total, current_index, gaps, num_elements


def extract_presentation_style(presentation: dict) -> str:
    """Extract dominant visual style patterns from the presentation.
    Includes white (#ffffff) and black (#000000) when the deck uses them, so title fills
    and body text color are preserved.
    """
    fonts: dict[str, int] = {}
    bg_colors: dict[str, int] = {}
    text_colors: dict[str, int] = {}
    border_colors: dict[str, int] = {}
    for slide in presentation.get("slides", []):
        for el in slide.get("pageElements", []):
            shape = el.get("shape", {})
            props = shape.get("shapeProperties", {})
            bg_fill = props.get("shapeBackgroundFill", {})
            solid = bg_fill.get("solidFill", {})
            bg_rgb = solid.get("color", {}).get("rgbColor", {})
            if bg_rgb:
                r = int(bg_rgb.get("red", 1) * 255)
                g = int(bg_rgb.get("green", 1) * 255)
                b = int(bg_rgb.get("blue", 1) * 255)
                h = f"#{r:02x}{g:02x}{b:02x}"
                bg_colors[h] = bg_colors.get(h, 0) + 1
            outline = props.get("outline", {})
            out_fill = outline.get("outlineFill", {}).get("solidFill", {}).get("color", {}).get("rgbColor", {})
            if out_fill:
                r = int(out_fill.get("red", 0) * 255)
                g = int(out_fill.get("green", 0) * 255)
                b = int(out_fill.get("blue", 0) * 255)
                h = f"#{r:02x}{g:02x}{b:02x}"
                border_colors[h] = border_colors.get(h, 0) + 1
            for te in shape.get("text", {}).get("textElements", []):
                run = te.get("textRun", {})
                style = run.get("style", {})
                if not run.get("content", "").strip():
                    continue
                font = style.get("fontFamily")
                if font:
                    fonts[font] = fonts.get(font, 0) + 1
                fg = style.get("foregroundColor", {}).get("opaqueColor", {}).get("rgbColor", {})
                if fg:
                    r = int(fg.get("red", 0) * 255)
                    g = int(fg.get("green", 0) * 255)
                    b = int(fg.get("blue", 0) * 255)
                    h = f"#{r:02x}{g:02x}{b:02x}"
                    text_colors[h] = text_colors.get(h, 0) + 1

    lines = ["Presentation visual style:"]
    top_fonts = sorted(fonts, key=fonts.get, reverse=True)[:5]
    if top_fonts:
        lines.append(f"  Common fonts: {', '.join(top_fonts)}")
    top_bg = sorted(bg_colors, key=bg_colors.get, reverse=True)[:6]
    if top_bg:
        lines.append(f"  Common background fills (use for title/body boxes): {', '.join(top_bg)}")
    top_tc = sorted(text_colors, key=text_colors.get, reverse=True)[:5]
    if top_tc:
        lines.append(f"  Common text colors: {', '.join(top_tc)}")
    top_border = sorted(border_colors, key=border_colors.get, reverse=True)[:4]
    if top_border:
        lines.append(f"  Common border/outline colors (use for border_color): {', '.join(top_border)}")
    # Hint for title vs list styling when deck uses white/grey/black
    if "#ffffff" in bg_colors or "#000000" in text_colors:
        lines.append("  Match element types: use the same fill and text color as similar elements on existing slides (e.g. title box = same fill as other title boxes, often #ffffff; list/body boxes = same fill as other list boxes, often light grey; text = same as deck, often #000000). Use shape_type RECTANGLE or TEXT_BOX with sharp corners to match.")
    return "\n".join(lines)


def _rgb_to_hex(rgb: dict[str, Any]) -> Optional[str]:
    """Convert Slides rgbColor dict to #RRGGBB."""
    if not isinstance(rgb, dict):
        return None
    if not {"red", "green", "blue"} & set(rgb.keys()):
        return None
    r = int(float(rgb.get("red", 0)) * 255)
    g = int(float(rgb.get("green", 0)) * 255)
    b = int(float(rgb.get("blue", 0)) * 255)
    return f"#{r:02x}{g:02x}{b:02x}"


def _build_theme_color_map(presentation: dict) -> dict[str, str]:
    """
    Build themeColor -> hex map from masters/layouts color schemes.
    Example keys: DARK1, LIGHT1, ACCENT1, ACCENT2, ...
    """
    theme_map: dict[str, str] = {}

    def _ingest_scope(scope: dict) -> None:
        # In Slides payloads, colorScheme is usually under pageProperties.colorScheme.
        scheme = (
            (scope.get("pageProperties", {}) or {}).get("colorScheme", {})
            or scope.get("colorScheme", {})
            or {}
        )
        for c in scheme.get("colors", []) or []:
            ctype = c.get("type")
            color = c.get("color", {}) or {}
            rgb_hex = _rgb_to_hex(color.get("rgbColor", {}) or {})
            if ctype and rgb_hex:
                theme_map[ctype] = rgb_hex

    for m in presentation.get("masters", []) or []:
        _ingest_scope(m)
    for l in presentation.get("layouts", []) or []:
        _ingest_scope(l)
    # Some decks may expose a top-level pageProperties color scheme as well.
    _ingest_scope(presentation)

    return theme_map


def _ingest_theme_from_scope(scope: dict, theme_map: dict[str, str]) -> None:
    """
    Merge theme colors from a page/presentation scope into an existing map.
    Useful when pages.get returns per-page color schemes.
    """
    scheme = (
        (scope.get("pageProperties", {}) or {}).get("colorScheme", {})
        or scope.get("colorScheme", {})
        or {}
    )
    for c in scheme.get("colors", []) or []:
        ctype = c.get("type")
        color = c.get("color", {}) or {}
        rgb_hex = _rgb_to_hex(color.get("rgbColor", {}) or {})
        if ctype and rgb_hex and ctype not in theme_map:
            theme_map[ctype] = rgb_hex


def _resolve_color_to_hex(color_obj: dict[str, Any], theme_map: dict[str, str], counters: dict) -> Optional[str]:
    """
    Resolve Slides color object to hex:
    - color.rgbColor
    - color.themeColor (mapped via theme color scheme)
    - foregroundColor.opaqueColor.{rgbColor|themeColor}
    """
    if not isinstance(color_obj, dict) or not color_obj:
        return None

    # Handle wrapper used by text foregroundColor
    if "opaqueColor" in color_obj and isinstance(color_obj["opaqueColor"], dict):
        color_obj = color_obj["opaqueColor"]

    rgb_hex = _rgb_to_hex(color_obj.get("rgbColor", {}) or {})
    if rgb_hex:
        counters["rgb_color_hits"] = counters.get("rgb_color_hits", 0) + 1
        return rgb_hex

    theme_key = color_obj.get("themeColor")
    if theme_key and theme_key in theme_map:
        counters["theme_color_hits"] = counters.get("theme_color_hits", 0) + 1
        return theme_map[theme_key]

    if theme_key:
        counters["theme_color_misses"] = counters.get("theme_color_misses", 0) + 1
        miss_keys = counters.setdefault("theme_color_miss_keys", {})
        miss_keys[theme_key] = miss_keys.get(theme_key, 0) + 1
    return None


def _extract_style_from_elements(
    page_el: list,
    fonts: dict,
    bg_colors: dict,
    text_colors: dict,
    border_colors: dict,
    theme_map: dict[str, str],
    counters: dict,
    debug_once: list,
    source_label: str,
) -> None:
    """Update style dicts from a list of page elements."""
    for el in page_el:
        shape = el.get("shape", {})
        if not shape:
            continue
        counters["shapes_seen"] = counters.get("shapes_seen", 0) + 1

        props = shape.get("shapeProperties", {})
        bg_fill = props.get("shapeBackgroundFill", {})
        solid = bg_fill.get("solidFill", {})
        bg_hex = _resolve_color_to_hex(solid.get("color", {}) or {}, theme_map, counters)
        if bg_hex:
            bg_colors[bg_hex] = bg_colors.get(bg_hex, 0) + 1

        outline = props.get("outline", {})
        out_fill = outline.get("outlineFill", {}).get("solidFill", {}).get("color", {})
        border_hex = _resolve_color_to_hex(out_fill or {}, theme_map, counters)
        if border_hex:
            border_colors[border_hex] = border_colors.get(border_hex, 0) + 1

        # Collect run styles first, then apply inheritance fallback for empty runs
        run_items: list[tuple[str, dict]] = []
        for te in shape.get("text", {}).get("textElements", []) or []:
            run = te.get("textRun", {})
            if not run:
                continue
            counters["text_runs_seen"] = counters.get("text_runs_seen", 0) + 1
            content = (run.get("content") or "")
            style = run.get("style") or run.get("textStyle") or {}
            if not style:
                # Fallback source for some payloads
                pm_style = te.get("paragraphMarker", {}).get("style", {}) or {}
                if pm_style:
                    style = pm_style
                    counters["paragraph_style_fallback_hits"] = counters.get("paragraph_style_fallback_hits", 0) + 1
            run_items.append((content, style))

            if not debug_once[0] and (content.strip() or style):
                fg = style.get("foregroundColor", {}) if isinstance(style, dict) else {}
                print(
                    f"   STYLE_VALUES: sample[{source_label}] textRun keys={list(run.keys())!r} "
                    f"style keys={list(style.keys()) if isinstance(style, dict) else type(style)!r}"
                )
                print(
                    f"   STYLE_VALUES: sample[{source_label}] foregroundColor keys="
                    f"{list(fg.keys()) if isinstance(fg, dict) else type(fg)!r}"
                )
                debug_once[0] = True

        # Pass 1: explicit styles
        shape_fonts: dict[str, int] = {}
        shape_colors: dict[str, int] = {}
        missing_text_runs = 0

        for content, style in run_items:
            if not content.strip():
                continue
            if not isinstance(style, dict) or not style:
                missing_text_runs += 1
                continue

            font = style.get("fontFamily") or (style.get("weightedFontFamily", {}) or {}).get("fontFamily")
            if font:
                fonts[font] = fonts.get(font, 0) + 1
                shape_fonts[font] = shape_fonts.get(font, 0) + 1
                counters["explicit_font_hits"] = counters.get("explicit_font_hits", 0) + 1

            text_hex = _resolve_color_to_hex(style.get("foregroundColor", {}) or {}, theme_map, counters)
            if text_hex:
                text_colors[text_hex] = text_colors.get(text_hex, 0) + 1
                shape_colors[text_hex] = shape_colors.get(text_hex, 0) + 1
                counters["explicit_text_color_hits"] = counters.get("explicit_text_color_hits", 0) + 1

        # Pass 2: inheritance fallback when runs have content but no explicit style
        if missing_text_runs > 0:
            fallback_font = None
            if shape_fonts:
                fallback_font = max(shape_fonts, key=shape_fonts.get)
            elif fonts:
                fallback_font = max(fonts, key=fonts.get)

            fallback_color = None
            if shape_colors:
                fallback_color = max(shape_colors, key=shape_colors.get)
            elif text_colors:
                fallback_color = max(text_colors, key=text_colors.get)

            if fallback_font:
                fonts[fallback_font] = fonts.get(fallback_font, 0) + missing_text_runs
                counters["inherited_font_fallback_hits"] = counters.get("inherited_font_fallback_hits", 0) + missing_text_runs
            if fallback_color:
                text_colors[fallback_color] = text_colors.get(fallback_color, 0) + missing_text_runs
                counters["inherited_text_color_fallback_hits"] = counters.get("inherited_text_color_fallback_hits", 0) + missing_text_runs


def get_presentation_style_values(
    presentation: dict,
    presentation_id: Optional[str] = None,
    access_token: Optional[str] = None,
) -> dict:
    """
    Return the deck's primary font, text color, background fills, and border colors.
    When presentation_id and access_token are provided, fetches each slide via pages.get
    so we use full page content (presentations.get may not include full pageElements).
    """
    fonts: dict[str, int] = {}
    bg_colors: dict[str, int] = {}
    text_colors: dict[str, int] = {}
    border_colors: dict[str, int] = {}
    debug_once = [False]
    counters: dict[str, int] = {}
    theme_map = _build_theme_color_map(presentation)
    print(f"   STYLE_VALUES: theme colors loaded={len(theme_map)} keys={list(theme_map.keys())[:8]}")

    slides = presentation.get("slides", [])
    if presentation_id and access_token and slides:
        # Fetch each slide page so we get full pageElements (correct font/color/outline)
        print(f"   STYLE_VALUES: fetching {len(slides)} slides via pages.get for full style")
        for slide in slides:
            sid = slide.get("objectId")
            if not sid:
                continue
            try:
                page = get_slide_elements(presentation_id, sid, access_token)
                _ingest_theme_from_scope(page, theme_map)
                page_el = page.get("pageElements", [])
                _extract_style_from_elements(
                    page_el, fonts, bg_colors, text_colors, border_colors,
                    theme_map, counters, debug_once, source_label=f"slide:{sid}"
                )
            except Exception as e:
                print(f"   STYLE_VALUES: pages.get for {sid} failed: {e}")

        # Also fetch layouts/masters via pages.get so inherited defaults are visible.
        for layout in presentation.get("layouts", []) or []:
            lid = layout.get("objectId")
            if not lid:
                continue
            try:
                page = get_slide_elements(presentation_id, lid, access_token)
                _ingest_theme_from_scope(page, theme_map)
                page_el = page.get("pageElements", [])
                _extract_style_from_elements(
                    page_el, fonts, bg_colors, text_colors, border_colors,
                    theme_map, counters, debug_once, source_label=f"layout-page:{lid}"
                )
            except Exception as e:
                print(f"   STYLE_VALUES: pages.get for layout {lid} failed: {e}")

        for master in presentation.get("masters", []) or []:
            mid = master.get("objectId")
            if not mid:
                continue
            try:
                page = get_slide_elements(presentation_id, mid, access_token)
                _ingest_theme_from_scope(page, theme_map)
                page_el = page.get("pageElements", [])
                _extract_style_from_elements(
                    page_el, fonts, bg_colors, text_colors, border_colors,
                    theme_map, counters, debug_once, source_label=f"master-page:{mid}"
                )
            except Exception as e:
                print(f"   STYLE_VALUES: pages.get for master {mid} failed: {e}")
    else:
        print(f"   STYLE_VALUES: scanning {len(slides)} slides from presentation (keys: {list(presentation.keys())})")
        for si, slide in enumerate(slides):
            page_el = slide.get("pageElements", [])
            if not page_el and slide:
                print(f"   STYLE_VALUES: slide[{si}] {slide.get('objectId', '?')!r} has no pageElements (keys: {list(slide.keys())})")
            _extract_style_from_elements(
                page_el, fonts, bg_colors, text_colors, border_colors,
                theme_map, counters, debug_once, source_label=f"presentation-slide:{si}"
            )

    # Also scan layouts/masters for inherited defaults (fonts/colors often live here)
    for li, layout in enumerate(presentation.get("layouts", []) or []):
        _extract_style_from_elements(
            layout.get("pageElements", []) or [],
            fonts, bg_colors, text_colors, border_colors,
            theme_map, counters, debug_once, source_label=f"layout:{li}"
        )
    for mi, master in enumerate(presentation.get("masters", []) or []):
        _extract_style_from_elements(
            master.get("pageElements", []) or [],
            fonts, bg_colors, text_colors, border_colors,
            theme_map, counters, debug_once, source_label=f"master:{mi}"
        )

    print(f"   STYLE_VALUES: theme colors after pages scan={len(theme_map)} keys={list(theme_map.keys())[:8]}")

    top_fonts = sorted(fonts, key=fonts.get, reverse=True)
    top_tc = sorted(text_colors, key=text_colors.get, reverse=True)
    top_bg = sorted(bg_colors, key=bg_colors.get, reverse=True)
    top_border = sorted(border_colors, key=border_colors.get, reverse=True)

    primary_font = top_fonts[0] if top_fonts else None
    primary_text_color = top_tc[0] if top_tc else None
    primary_background_fills = top_bg[:6]
    primary_border_colors = top_border[:4]

    print(f"   STYLE_VALUES: primary_font={primary_font!r} primary_text_color={primary_text_color!r}")
    print(f"   STYLE_VALUES: primary_background_fills={primary_background_fills!r} primary_border_colors={primary_border_colors!r}")
    print(f"   STYLE_VALUES: counters={counters}")
    if not primary_font and fonts:
        print(f"   STYLE_VALUES: WARNING fonts dict had {len(fonts)} entries but top_fonts empty")
    if not primary_text_color and text_colors:
        print(f"   STYLE_VALUES: WARNING text_colors had {len(text_colors)} entries but top_tc empty")

    return {
        "primary_font": primary_font,
        "primary_text_color": primary_text_color,
        "primary_background_fills": primary_background_fills,
        "primary_border_colors": primary_border_colors,
    }
