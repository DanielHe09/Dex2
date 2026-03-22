"""
Layout: two-phase content-then-layout (role-based) and post-process to fix overlaps/alignment.
"""

from __future__ import annotations

import re
from typing import Any

# Defaults for layout
MARGIN_PT = 36
TITLE_Y_PT = 24
TITLE_HEIGHT_PT = 48
ROW_HEIGHT_PT = 56
NUMBER_BOX_WIDTH_PT = 40
NUMBER_TEXT_GAP_PT = 12
MIN_TEXT_BOX_HEIGHT_PT = 32
# Empty "horizontal" text boxes in whitespace often come back ~10–24pt tall; expand those.
MIN_USABLE_EMPTY_HORIZONTAL_TEXT_BOX_HEIGHT_PT = 72
MAX_USABLE_EMPTY_HORIZONTAL_TEXT_BOX_HEIGHT_PT = 120
WIDE_EMPTY_TEXT_BOX_MIN_WIDTH_RATIO = 0.4  # of content width
SLIVER_MAX_HEIGHT_PT = 52  # if wider than ratio and shorter than this, expand
SMALL_AREA_THRESHOLD_PT2 = 4000  # below this = likely number box
OVERLAP_GAP_PT = 8


def _bbox(inst: dict) -> tuple[float, float, float, float] | None:
    """Return (x, y, w, h) for a create_shape instruction, or None if missing."""
    if inst.get("action") != "create_shape":
        return None
    x = inst.get("x_pt")
    y = inst.get("y_pt")
    w = inst.get("width_pt")
    h = inst.get("height_pt")
    if x is None or y is None or w is None or h is None:
        return None
    return (float(x), float(y), float(w), float(h))


def _set_bbox(inst: dict, x_pt: float, y_pt: float, width_pt: float, height_pt: float) -> None:
    """Set position/size on a create_shape instruction in place."""
    inst["x_pt"] = x_pt
    inst["y_pt"] = y_pt
    inst["width_pt"] = width_pt
    inst["height_pt"] = height_pt


def _area(w: float, h: float) -> float:
    return w * h


def fix_layout(
    instructions: list[dict],
    page_width_pt: float,
    page_height_pt: float,
) -> list[dict]:
    """
    Post-process create_shape instructions: detect number+text pairs, align side-by-side,
    and resolve overlaps. Other instructions (create_line, move, etc.) are left unchanged.
    """
    # Collect create_shape with valid bbox and their indices
    shape_indices: list[int] = []
    for i, inst in enumerate(instructions):
        if inst.get("action") == "create_shape" and _bbox(inst) is not None:
            shape_indices.append(i)

    if not shape_indices:
        return instructions

    # Work with copies so we don't mutate originals until we're sure
    out = [dict(inst) for inst in instructions]

    # Build list of (index_in_out, x, y, w, h) for create_shape only
    boxes: list[tuple[int, float, float, float, float]] = []
    for i in shape_indices:
        b = _bbox(out[i])
        if b:
            boxes.append((i, b[0], b[1], b[2], b[3]))

    # Sort by y then x so we process top-to-bottom, left-to-right
    boxes.sort(key=lambda t: (t[2], t[1]))

    # Detect number+text pairs: consecutive pairs where first has small area, second has larger
    used = set()
    pairs: list[tuple[int, int]] = []  # (idx_small, idx_large) into boxes
    for j in range(len(boxes) - 1):
        if j in used:
            continue
        idx_a, xa, ya, wa, ha = boxes[j]
        idx_b, xb, yb, wb, hb = boxes[j + 1]
        if (j + 1) in used:
            continue
        # Same row: y overlap or very close
        if abs((ya + ha / 2) - (yb + hb / 2)) > max(ha, hb) * 0.6:
            continue
        area_a = _area(wa, ha)
        area_b = _area(wb, hb)
        if area_a < SMALL_AREA_THRESHOLD_PT2 and area_b > area_a:
            pairs.append((j, j + 1))
            used.add(j)
            used.add(j + 1)

    # Apply pair layout: number left, text right, same baseline
    right_edge = page_width_pt - MARGIN_PT
    num_width = NUMBER_BOX_WIDTH_PT
    gap = NUMBER_TEXT_GAP_PT
    text_start = MARGIN_PT + num_width + gap
    text_width = right_edge - text_start

    for (j_small, j_large) in pairs:
        i_small = boxes[j_small][0]
        i_large = boxes[j_large][0]
        _, x1, y1, w1, h1 = boxes[j_small]
        _, x2, y2, w2, h2 = boxes[j_large]
        # Use the row y as the smaller one's y; align the larger to same row
        row_y = min(y1, y2)
        # Number box: left, fixed width
        _set_bbox(out[i_small], MARGIN_PT, row_y, num_width, max(h1, MIN_TEXT_BOX_HEIGHT_PT))
        # Text box: right of number
        _set_bbox(out[i_large], text_start, row_y, text_width, max(h2, MIN_TEXT_BOX_HEIGHT_PT))

    # For any create_shape not in a pair, check overlap with others and shift if needed
    for idx in shape_indices:
        b = _bbox(out[idx])
        if b is None:
            continue
        x, y, w, h = b
        # Collect other bboxes (after our layout changes)
        others = []
        for i in shape_indices:
            if i == idx:
                continue
            ob = _bbox(out[i])
            if ob:
                others.append(ob)
        # If we overlap any other, shift down
        for (ox, oy, ow, oh) in others:
            if _overlap(x, y, w, h, ox, oy, ow, oh):
                # shift current down below (oy + oh)
                new_y = oy + oh + OVERLAP_GAP_PT
                if new_y + h <= page_height_pt:
                    _set_bbox(out[idx], x, new_y, w, h)
                    # update our bbox for subsequent checks
                    x, y, w, h = x, new_y, w, h

    return out


def _overlap(
    x1: float, y1: float, w1: float, h1: float,
    x2: float, y2: float, w2: float, h2: float,
) -> bool:
    """True if the two axis-aligned rectangles overlap."""
    return not (x1 + w1 <= x2 or x2 + w2 <= x1 or y1 + h1 <= y2 or y2 + h2 <= y1)


def _has_role(inst: dict) -> bool:
    return bool(inst.get("role")) and inst.get("action") == "create_shape"


def _parse_role(role: str) -> tuple[str, int | None]:
    """Parse role like 'title', 'item_1_number', 'item_2_text'. Returns (kind, item_index)."""
    role = (role or "").strip().lower()
    if role == "title":
        return ("title", None)
    m = re.match(r"item_(\d+)_(number|text)", role)
    if m:
        return ("item", int(m.group(1)))
    return ("other", None)


def compute_layout_from_roles(
    instructions: list[dict],
    page_width_pt: float,
    page_height_pt: float,
) -> list[dict]:
    """
    Two-phase layout: instructions with "role" (and no x_pt) get positions assigned here.
    Supports: title, item_N_number, item_N_text. Other instructions are returned unchanged
    (or with default position if create_shape without role).
    """
    # Collect role-based create_shape (mutate in place)
    role_specs: list[dict] = []
    for inst in instructions:
        if _has_role(inst) and inst.get("x_pt") is None and inst.get("y_pt") is None:
            role_specs.append(inst)

    if not role_specs:
        return instructions

    # Group by role: title(s), then item 1 number/text, item 2 number/text, ...
    titles: list[dict] = []
    items: dict[int, list[dict]] = {}  # item_index -> [number_spec, text_spec] or [text_only]
    for spec in role_specs:
        kind, idx = _parse_role(spec.get("role", ""))
        if kind == "title":
            titles.append(spec)
        elif kind == "item" and idx is not None:
            if idx not in items:
                items[idx] = []
            items[idx].append(spec)

    # Assign positions
    y_cursor = TITLE_Y_PT
    margin = MARGIN_PT
    right_edge = page_width_pt - margin
    num_width = NUMBER_BOX_WIDTH_PT
    gap = NUMBER_TEXT_GAP_PT
    text_start = margin + num_width + gap
    text_width = right_edge - text_start

    for t in titles:
        _set_bbox(t, margin, y_cursor, page_width_pt - 2 * margin, TITLE_HEIGHT_PT)
        y_cursor += TITLE_HEIGHT_PT + 16

    for idx in sorted(items.keys()):
        row_specs = items[idx]
        # Determine if we have number + text or just text
        roles_in_row = [s.get("role", "") for s in row_specs]
        has_number = any("number" in r for r in roles_in_row)
        has_text = any("text" in r for r in roles_in_row)
        row_h = ROW_HEIGHT_PT
        for spec in row_specs:
            r = (spec.get("role") or "").lower()
            if "number" in r:
                _set_bbox(spec, margin, y_cursor, num_width, row_h)
            elif "text" in r:
                _set_bbox(spec, text_start, y_cursor, text_width, row_h)
            else:
                _set_bbox(spec, text_start, y_cursor, text_width, row_h)
        y_cursor += row_h + 8

    # Fallback: any role spec without coords (e.g. unrecognized role) gets a default row
    for spec in role_specs:
        if spec.get("x_pt") is None:
            _set_bbox(spec, margin, y_cursor, text_width, ROW_HEIGHT_PT)
            y_cursor += ROW_HEIGHT_PT + 8

    return instructions


def expand_sliver_empty_text_boxes(
    instructions: list[dict],
    page_width_pt: float,
    page_height_pt: float,
) -> list[dict]:
    """
    Widen/tall-en empty TEXT_BOX shapes that are wide but extremely short (vision/GPT
    often returns a 1-line-height box for "empty text box in whitespace").
    """
    content_w = page_width_pt - 2 * MARGIN_PT
    out: list[dict] = []
    for inst in instructions:
        inst = dict(inst)
        if inst.get("action") != "create_shape":
            out.append(inst)
            continue
        st = (inst.get("shape_type") or "TEXT_BOX").upper()
        if st != "TEXT_BOX":
            out.append(inst)
            continue
        if (inst.get("text") or "").strip():
            out.append(inst)
            continue
        x = float(inst.get("x_pt") or 0)
        y = float(inst.get("y_pt") or 0)
        w = float(inst.get("width_pt") or 0)
        h = float(inst.get("height_pt") or 0)
        if w >= content_w * WIDE_EMPTY_TEXT_BOX_MIN_WIDTH_RATIO and h <= SLIVER_MAX_HEIGHT_PT:
            room_below = page_height_pt - y - MARGIN_PT
            target_h = max(
                MIN_USABLE_EMPTY_HORIZONTAL_TEXT_BOX_HEIGHT_PT,
                min(MAX_USABLE_EMPTY_HORIZONTAL_TEXT_BOX_HEIGHT_PT, room_below * 0.78),
            )
            new_h = min(target_h, max(MIN_TEXT_BOX_HEIGHT_PT, room_below - 4))
            if new_h > h:
                print(
                    f"   LAYOUT: expanded empty horizontal TEXT_BOX "
                    f"height {h:.1f} -> {new_h:.1f} pt (whitespace strip -> usable area)"
                )
                inst["height_pt"] = round(new_h, 1)
            # Do NOT change x_pt / width_pt here — forcing full-bleed width breaks alignment
            # with column grids below and overrides model/vision border choices.
        out.append(inst)
    return out


def prepare_instructions_for_apply(
    instructions: list[dict],
    page_width_pt: float,
    page_height_pt: float,
) -> list[dict]:
    """
    Run two-phase (role-based) then post-process. Call this before apply_instructions.
    - If any create_shape has "role" and no x_pt, run compute_layout_from_roles.
    - Then run fix_layout on all create_shape instructions that have coordinates.
    """
    if not instructions:
        return instructions

    has_roles = any(_has_role(inst) and inst.get("x_pt") is None for inst in instructions)
    if has_roles:
        instructions = compute_layout_from_roles(instructions, page_width_pt, page_height_pt)

    fixed = fix_layout(instructions, page_width_pt, page_height_pt)
    return expand_sliver_empty_text_boxes(fixed, page_width_pt, page_height_pt)


# Fallbacks when the deck has no style data (force consistent black text, Arial, white fill, black border)
DEFAULT_TEXT_COLOR = "#000000"
DEFAULT_FONT = "Arial"
DEFAULT_FILL = "#ffffff"
DEFAULT_BORDER_COLOR = "#000000"
DEFAULT_BORDER_WEIGHT_PT = 1


def normalize_instructions_style(
    instructions: list[dict],
    style_values: dict,
    fill_missing_only: bool = False,
) -> list[dict]:
    """
    Force text color, font, fill, and outline on every create_shape to match the deck
    (or fallbacks) before batch update. When fill_missing_only is True, only set a field
    if the instruction does not already have it (preserves vision-generated style).
    """
    primary_font = style_values.get("primary_font") or DEFAULT_FONT
    primary_text_color = style_values.get("primary_text_color") or DEFAULT_TEXT_COLOR
    bg_fills = style_values.get("primary_background_fills") or [DEFAULT_FILL]
    border_colors = style_values.get("primary_border_colors") or [DEFAULT_BORDER_COLOR]

    if fill_missing_only:
        print(
            f"   NORMALIZE_STYLE: fill_missing_only=True — bundle for gaps only: "
            f"font={primary_font!r} text={primary_text_color!r} fill={bg_fills[0]!r} "
            f"border_fallback={border_colors[0]!r} "
            f"(instruction keeps border/font/color if already set; see per-shape line)"
        )
    else:
        print(
            f"   NORMALIZE_STYLE: forcing every field from bundle: font={primary_font!r} "
            f"text_color={primary_text_color!r} fill={bg_fills[0]!r} border={border_colors[0]!r}"
        )

    out = []
    for i, inst in enumerate(instructions):
        inst = dict(inst)
        if inst.get("action") == "create_shape":
            if not fill_missing_only or inst.get("color") is None:
                inst["color"] = primary_text_color
            if not fill_missing_only or inst.get("font_family") is None:
                inst["font_family"] = primary_font
            role = (inst.get("role") or "").lower()
            if not fill_missing_only or inst.get("background_color") is None:
                if role == "title" and len(bg_fills) >= 1:
                    inst["background_color"] = bg_fills[0]
                elif "item" in role and len(bg_fills) >= 2:
                    inst["background_color"] = bg_fills[1]
                else:
                    inst["background_color"] = bg_fills[0]
            if not fill_missing_only or inst.get("border_color") is None:
                inst["border_color"] = border_colors[0]
            if inst.get("border_weight_pt") is None:
                inst["border_weight_pt"] = DEFAULT_BORDER_WEIGHT_PT
            print(f"   NORMALIZE_STYLE: create_shape[{i}] role={role!r} -> color={inst['color']!r} font_family={inst['font_family']!r} bg={inst['background_color']!r} border={inst['border_color']!r}")
        out.append(inst)
    return out
