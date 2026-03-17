"""
Text Formatter - XML-level paragraph formatting operations.

This module encapsulates all direct XML manipulation for paragraph properties (pPr),
bullet types, and style resolution. It's the surgical fix for the root cause:
`tf.add_paragraph()` creates bare `<a:p>` elements that lose template formatting.

By deep-copying the first paragraph's pPr XML, every subsequent paragraph matches
the template designer's intent.

All functions operate on raw lxml elements — no knowledge of slides, layouts, or roles.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import Any, Optional

from lxml import etree
from pptx.oxml.ns import qn

# OOXML element ordering within a:pPr (from ECMA-376 spec)
# These elements MUST appear in this exact order for valid XML
_PPR_CHILD_ORDER = [
    "lnSpc",
    "spcBef",
    "spcAft",
    "buClrTx",
    "buClr",
    "buSzTx",
    "buSzPts",
    "buSzPct",
    "buFontTx",
    "buFont",
    "buNone",
    "buAutoNum",
    "buChar",
    "buBlip",
    "tabLst",
    "defRPr",
    "extLst",
]


def _get_tag_order(tag_name: str) -> int:
    """Get the canonical position index for a pPr child element tag.

    Args:
        tag_name: Local tag name (e.g., 'buChar', 'defRPr')

    Returns:
        Position index (lower = earlier in sequence), or 999 if unknown
    """
    try:
        return _PPR_CHILD_ORDER.index(tag_name)
    except ValueError:
        return 999  # Unknown tags go at end


def insert_pPr_child_ordered(pPr_elem: etree._Element, child: etree._Element) -> None:
    """Insert a child element into a:pPr at the correct position per OOXML spec.

    The OOXML spec requires strict child element ordering within a:pPr.
    This function ensures the child is inserted at the correct position.

    Args:
        pPr_elem: The a:pPr element to insert into
        child: The child element to insert
    """
    child_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
    child_order = _get_tag_order(child_tag)

    # Find the first existing child that comes AFTER this element in the canonical order
    insert_pos = len(pPr_elem)  # Default: append at end
    for i, existing in enumerate(pPr_elem):
        existing_tag = existing.tag.split("}")[-1] if "}" in existing.tag else existing.tag
        existing_order = _get_tag_order(existing_tag)
        if existing_order > child_order:
            insert_pos = i
            break

    pPr_elem.insert(insert_pos, child)


def ensure_pPr(p_elem: etree._Element) -> etree._Element:
    """Ensure paragraph has a:pPr element, creating one if needed.

    a:pPr must be the FIRST child of a:p (before any a:r elements).

    Args:
        p_elem: The a:p paragraph element

    Returns:
        The a:pPr element (existing or newly created)
    """
    pPr = p_elem.find(qn("a:pPr"))
    if pPr is None:
        pPr = etree.Element(qn("a:pPr"))
        # Insert at index 0 (before any run elements)
        p_elem.insert(0, pPr)
    return pPr


def copy_paragraph_properties(source_pPr: Optional[etree._Element], target_p: etree._Element) -> None:
    """Deep-copy a:pPr from source to target paragraph element.

    Copies all paragraph properties: marL, indent, algn, lnSpc, spcBef, spcAft,
    bullet elements, defRPr (font size, bold, italic, color, font family).

    Does NOT copy: lvl attribute (caller sets p.level separately via python-pptx API).

    Strategy: Deep-copy the entire a:pPr element, remove lvl attribute from copy,
    then replace any existing a:pPr on target or insert as first child.

    Args:
        source_pPr: The a:pPr element to copy from (can be None)
        target_p: The a:p paragraph element to copy to
    """
    if source_pPr is None:
        return

    # Deep-copy to avoid layout contamination (Pitfall 6)
    pPr_copy = copy.deepcopy(source_pPr)

    # Remove lvl attribute from copy (caller sets p.level separately)
    if "lvl" in pPr_copy.attrib:
        del pPr_copy.attrib["lvl"]

    # Remove any existing pPr on target
    existing_pPr = target_p.find(qn("a:pPr"))
    if existing_pPr is not None:
        target_p.remove(existing_pPr)

    # Insert copied pPr as first child (before any a:r elements)
    target_p.insert(0, pPr_copy)


def copy_run_properties(source_run_elem: etree._Element, target_run_elem: etree._Element) -> None:
    """Deep-copy a:rPr (run properties) from source run to target run.

    Propagates template's first run formatting (font name, size, bold, italic, color)
    to runs in added paragraphs.

    Args:
        source_run_elem: The a:r element to copy from
        target_run_elem: The a:r element to copy to
    """
    source_rPr = source_run_elem.find(qn("a:rPr"))
    if source_rPr is None:
        return  # Source has no run properties

    # Deep-copy run properties
    rPr_copy = copy.deepcopy(source_rPr)

    # Remove any existing rPr on target
    existing_rPr = target_run_elem.find(qn("a:rPr"))
    if existing_rPr is not None:
        target_run_elem.remove(existing_rPr)

    # Insert copied rPr as first child
    target_run_elem.insert(0, rPr_copy)


@dataclass
class BulletSpec:
    """Agent-facing bullet configuration."""

    type: str = "auto"  # "auto" | "bullet" | "number" | "none"
    char: str = ""  # Custom bullet char (e.g., "–", "✓") — only if type="bullet"
    start_at: int = 1  # Starting number — only if type="number"
    scheme: str = ""  # OOXML numbering scheme: "arabicPeriod", "alphaLcPeriod", etc.


def apply_bullet_type(p_elem: etree._Element, spec: BulletSpec, template_pPr: Optional[etree._Element] = None) -> None:
    """Set bullet type on a paragraph element based on BulletSpec.

    Strategy:
    1. If spec.type == "auto": copy bullet from template_pPr (or leave as-is)
    2. If spec.type == "bullet": set <a:buChar char="..."/>
    3. If spec.type == "number": set <a:buAutoNum type="..." startAt="..."/>
    4. If spec.type == "none": set <a:buNone/>

    Always removes existing bullet elements (buChar, buAutoNum, buNone, buBlip)
    before applying new one - they are mutually exclusive (Pitfall 7).

    Args:
        p_elem: The a:p paragraph element
        spec: Bullet specification
        template_pPr: Optional template pPr to copy bullets from when type="auto"
    """
    pPr = ensure_pPr(p_elem)

    # Bullet type elements are mutually exclusive - remove all existing
    BULLET_TYPE_TAGS = {"buNone", "buAutoNum", "buChar", "buBlip"}
    # Also remove bullet decoration elements when changing type
    BULLET_DECORATION_TAGS = {"buClrTx", "buClr", "buSzTx", "buSzPts", "buSzPct", "buFontTx", "buFont"}

    for child in list(pPr):
        local_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local_tag in BULLET_TYPE_TAGS or local_tag in BULLET_DECORATION_TAGS:
            pPr.remove(child)

    # Apply new bullet type
    if spec.type == "auto":
        # Copy bullet elements from template if provided
        if template_pPr is not None:
            for child in template_pPr:
                local_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if local_tag in BULLET_TYPE_TAGS or local_tag in BULLET_DECORATION_TAGS:
                    child_copy = copy.deepcopy(child)
                    insert_pPr_child_ordered(pPr, child_copy)

    elif spec.type == "bullet":
        # Create buChar element
        buChar = etree.Element(qn("a:buChar"))
        # Use explicit char if provided, else try template, else default to "•"
        if spec.char:
            buChar.set("char", spec.char)
        elif template_pPr is not None:
            template_buChar = template_pPr.find(qn("a:buChar"))
            if template_buChar is not None:
                buChar.set("char", template_buChar.get("char", "•"))
            else:
                buChar.set("char", "•")
        else:
            buChar.set("char", "•")
        insert_pPr_child_ordered(pPr, buChar)

    elif spec.type == "number":
        # Create buAutoNum element
        buAutoNum = etree.Element(qn("a:buAutoNum"))
        scheme = spec.scheme if spec.scheme else "arabicPeriod"
        buAutoNum.set("type", scheme)
        if spec.start_at != 1:
            buAutoNum.set("startAt", str(spec.start_at))
        insert_pPr_child_ordered(pPr, buAutoNum)

    elif spec.type == "none":
        # Create buNone element
        buNone = etree.Element(qn("a:buNone"))
        insert_pPr_child_ordered(pPr, buNone)


def resolve_level_style(placeholder_elem, level: int) -> dict[str, Any]:
    """Walk XML inheritance chain to find effective text styles for a specific indent level.

    Resolution order:
    1. Placeholder's own txBody > lstStyle > lvl{N}pPr
    2. If not found, return empty dict (caller can fall back to other resolution methods)

    Returns dict with keys: indent (marL), bullet_indent (indent attr), line_spacing,
    space_before, space_after, font_size, bold, italic, bullet_type, bullet_char.

    Note: spcBef and spcAft in OOXML are CHILD ELEMENTS, not attributes.
    Correct XML: <a:spcBef><a:spcPts val="600"/></a:spcBef> where val is in hundredths of a point.

    Args:
        placeholder_elem: python-pptx placeholder object
        level: Indent level (0-indexed)

    Returns:
        Dict of resolved style properties (empty if not found)
    """
    elem = placeholder_elem._element
    styles: dict[str, Any] = {}

    # Check placeholder's own txBody > lstStyle
    txBody = elem.find(qn("p:txBody"))
    if txBody is None:
        txBody = elem.find(qn("a:txBody"))

    if txBody is None:
        return styles

    lstStyle = txBody.find(qn("a:lstStyle"))
    if lstStyle is None:
        return styles

    # Find the level-specific pPr element (lvl0pPr = level 0, lvl1pPr = level 1, etc.)
    # Note: lstStyle uses lvl1pPr for level 0, lvl2pPr for level 1, etc. (1-indexed names)
    lvl_elem = lstStyle.find(qn(f"a:lvl{level + 1}pPr"))
    if lvl_elem is None:
        return styles

    # Extract indentation
    mar_l = lvl_elem.get("marL")
    if mar_l is not None:
        styles["indent"] = int(mar_l)

    indent_attr = lvl_elem.get("indent")
    if indent_attr is not None:
        styles["bullet_indent"] = int(indent_attr)

    # Extract line spacing (from child element)
    lnSpc = lvl_elem.find(qn("a:lnSpc"))
    if lnSpc is not None:
        spcPct = lnSpc.find(qn("a:spcPct"))
        if spcPct is not None:
            val = spcPct.get("val")
            if val:
                styles["line_spacing"] = int(val)

    # Extract space before (from child element)
    spcBef = lvl_elem.find(qn("a:spcBef"))
    if spcBef is not None:
        spcPts = spcBef.find(qn("a:spcPts"))
        if spcPts is not None:
            val = spcPts.get("val")
            if val:
                # val is in hundredths of a point (600 = 6pt)
                styles["space_before"] = int(val) / 100.0
        # Also support percentage variant
        spcPct = spcBef.find(qn("a:spcPct"))
        if spcPct is not None:
            val = spcPct.get("val")
            if val:
                styles["space_before_pct"] = int(val)

    # Extract space after (from child element)
    spcAft = lvl_elem.find(qn("a:spcAft"))
    if spcAft is not None:
        spcPts = spcAft.find(qn("a:spcPts"))
        if spcPts is not None:
            val = spcPts.get("val")
            if val:
                # val is in hundredths of a point (300 = 3pt)
                styles["space_after"] = int(val) / 100.0
        # Also support percentage variant
        spcPct = spcAft.find(qn("a:spcPct"))
        if spcPct is not None:
            val = spcPct.get("val")
            if val:
                styles["space_after_pct"] = int(val)

    # Extract text run properties (font size, bold, italic)
    defRPr = lvl_elem.find(qn("a:defRPr"))
    if defRPr is not None:
        sz = defRPr.get("sz")
        if sz:
            styles["font_size"] = int(sz) / 100  # Convert to points

        bold = defRPr.get("b")
        if bold is not None:
            styles["bold"] = bold == "1"

        italic = defRPr.get("i")
        if italic is not None:
            styles["italic"] = italic == "1"

    # Extract bullet type
    buChar = lvl_elem.find(qn("a:buChar"))
    if buChar is not None:
        styles["bullet_type"] = "char"
        styles["bullet_char"] = buChar.get("char", "•")

    buAutoNum = lvl_elem.find(qn("a:buAutoNum"))
    if buAutoNum is not None:
        styles["bullet_type"] = "autonum"

    buNone = lvl_elem.find(qn("a:buNone"))
    if buNone is not None:
        styles["bullet_type"] = "none"

    return styles
