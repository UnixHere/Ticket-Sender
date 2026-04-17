"""
SVG Template Preparation Helper (Figma-compatible)
----------------------------------------------------
Handles Figma-exported SVGs correctly:
  - Reads text inside <tspan> children (Figma's default structure)
  - Handles default-namespace SVGs without double-counting elements
  - Replaces text in both <text> and <tspan> nodes
  - Lets you nudge placeholder text left/right with an X offset
  - Extends the SVG canvas width and adds a proper QR sidebar
    (sidebar rect overlaps ticket by 1 px — eliminates the hairline gap)

Usage:
    python prepare_svg.py your_ticket.svg
"""

import sys
import re
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------------------
# Namespace helpers
# ---------------------------------------------------------------------------

def _strip_ns(tag: str) -> str:
    """Return tag name without any {namespace} prefix."""
    return tag.split('}', 1)[-1] if '}' in tag else tag


def _detect_namespace(root) -> str:
    """Return the SVG namespace URI found on the root element, or ''."""
    if root.tag.startswith('{'):
        return root.tag.split('}', 1)[0].lstrip('{')
    return ''


def _find_all(root, local_name: str):
    """Find all descendants whose local tag name matches, regardless of namespace."""
    results = []
    for elem in root.iter():
        if _strip_ns(elem.tag) == local_name:
            results.append(elem)
    return results


# ---------------------------------------------------------------------------
# Text content helpers
# ---------------------------------------------------------------------------

def _get_text_content(text_elem) -> str:
    """
    Get the visible string from a <text> element.
    Figma places actual content inside <tspan> children, not in text.text.
    """
    parts = []
    if text_elem.text and text_elem.text.strip():
        parts.append(text_elem.text.strip())
    for child in text_elem:
        if _strip_ns(child.tag) == 'tspan':
            t = child.text or ''
            if t.strip():
                parts.append(t.strip())
            # nested tspan
            for grandchild in child:
                if _strip_ns(grandchild.tag) == 'tspan':
                    t2 = grandchild.text or ''
                    if t2.strip():
                        parts.append(t2.strip())
    return ' '.join(parts) if parts else ''


def _set_text_content(text_elem, value: str):
    """
    Replace text in a <text> element and all its <tspan> descendants.
    Works whether Figma used tspan children or plain text.text.
    """
    tspans = [c for c in text_elem if _strip_ns(c.tag) == 'tspan']

    if tspans:
        # Keep first tspan, clear the rest
        first = tspans[0]
        # Handle nested tspan inside first tspan (Figma sometimes does this)
        inner = [c for c in first if _strip_ns(c.tag) == 'tspan']
        if inner:
            inner[0].text = value
            for extra in inner[1:]:
                extra.text = ''
        else:
            first.text = value

        # Blank out any sibling tspans (e.g. multi-line placeholders)
        for sibling in tspans[1:]:
            sibling.text = ''
            for gc in sibling:
                if _strip_ns(gc.tag) == 'tspan':
                    gc.text = ''

        text_elem.text = None  # clear any bare text node
    else:
        text_elem.text = value


def _shift_text_x(text_elem, delta: float):
    """
    Shift a <text> element (and its <tspan> children that carry an x attribute)
    horizontally by `delta` pixels.  Positive = right, negative = left.
    """
    if delta == 0:
        return

    def _shift_attr(elem, attr='x'):
        raw = elem.get(attr)
        if raw is None:
            return
        # x can be a single number or a space-separated list (for per-glyph positioning)
        parts = raw.strip().split()
        try:
            shifted = ' '.join(f'{float(p) + delta:.4f}' for p in parts)
            elem.set(attr, shifted)
        except ValueError:
            pass  # not a numeric value — leave it alone

    _shift_attr(text_elem, 'x')

    for child in text_elem:
        if _strip_ns(child.tag) == 'tspan':
            _shift_attr(child, 'x')
            for grandchild in child:
                if _strip_ns(grandchild.tag) == 'tspan':
                    _shift_attr(grandchild, 'x')


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def analyze_svg(filepath):
    """Parse SVG and show all text elements with their content."""
    print(f"\n📋 Analyzing {filepath}...\n")

    # Register common namespaces so they survive a write round-trip
    ET.register_namespace('', 'http://www.w3.org/2000/svg')
    ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')

    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
    except Exception as e:
        print(f"❌ Error reading SVG: {e}")
        return None, None

    ns = _detect_namespace(root)
    if ns:
        print(f"   (namespace detected: {ns})")

    text_elements = _find_all(root, 'text')

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for el in text_elements:
        eid = id(el)
        if eid not in seen:
            seen.add(eid)
            unique.append(el)
    text_elements = unique

    if not text_elements:
        print("⚠️  No <text> elements found.\n")
        print("   Possible reasons:")
        print("   • Text was converted to outlines/paths in Figma")
        print("     → Select text in Figma, DON'T use 'Outline stroke', export as SVG")
        print("   • The file uses <foreignObject> (rare in Figma exports)")
        print()
    else:
        print(f"Found {len(text_elements)} text element(s):\n")
        for i, text in enumerate(text_elements, 1):
            content = _get_text_content(text)
            elem_id = text.get('id', '')
            id_hint = f"  [id={elem_id}]" if elem_id else ''
            print(f"  {i}. \"{content}\"{id_hint}")

    return tree, root, text_elements


# ---------------------------------------------------------------------------
# Placeholder insertion  (with optional X nudge)
# ---------------------------------------------------------------------------

def _ask_offset(label: str) -> float:
    """Ask the user for an X offset in pixels for a placeholder. Returns 0.0 to skip."""
    raw = input(
        f"  X offset for {label} (pixels; + = right, – = left, Enter = no change): "
    ).strip()
    if not raw:
        return 0.0
    try:
        return float(raw)
    except ValueError:
        print("  ⚠  Not a number — using 0.")
        return 0.0


def add_placeholders(text_elements, name_index=None, class_index=None):
    if name_index is not None:
        elem = text_elements[name_index - 1]
        _set_text_content(elem, '{NAME_PLACEHOLDER}')
        print(f"✓ Element {name_index} → NAME placeholder")
        dx = _ask_offset("NAME")
        if dx:
            _shift_text_x(elem, dx)
            print(f"  → shifted {dx:+.1f} px")

    if class_index is not None:
        elem = text_elements[class_index - 1]
        _set_text_content(elem, '{CLASS_PLACEHOLDER}')
        print(f"✓ Element {class_index} → CLASS placeholder")
        dx = _ask_offset("CLASS")
        if dx:
            _shift_text_x(elem, dx)
            print(f"  → shifted {dx:+.1f} px")


# ---------------------------------------------------------------------------
# QR sidebar
# ---------------------------------------------------------------------------

def clip_original_content(root, orig_w, orig_h):
    """
    Wrap all existing SVG children in a <g> with a clipPath that strictly
    crops them to the original ticket dimensions.  This hides any invisible
    stray elements that sit outside the visible area and would otherwise
    create a gap between the ticket and the QR sidebar.
    """
    svg_ns = _detect_namespace(root)
    tag = lambda name: f'{{{svg_ns}}}{name}' if svg_ns else name

    # Build <clipPath id="ticket_clip"><rect .../></clipPath>
    defs = None
    for child in root:
        if _strip_ns(child.tag) == 'defs':
            defs = child
            break
    if defs is None:
        defs = ET.Element(tag('defs'))
        root.insert(0, defs)

    clip = ET.SubElement(defs, tag('clipPath'))
    clip.set('id', 'ticket_clip')
    rect = ET.SubElement(clip, tag('rect'))
    rect.set('x', '0')
    rect.set('y', '0')
    rect.set('width', str(orig_w))
    rect.set('height', str(orig_h))

    # Gather every child that is NOT <defs>
    children_to_wrap = [c for c in list(root) if _strip_ns(c.tag) != 'defs']

    # Remove them from root
    for c in children_to_wrap:
        root.remove(c)

    # Re-add inside a clipping group
    wrapper = ET.SubElement(root, tag('g'))
    wrapper.set('clip-path', 'url(#ticket_clip)')
    for c in children_to_wrap:
        wrapper.append(c)

    print(f"✓ Original content clipped to {orig_w:.0f} × {orig_h:.0f} (hides out-of-bounds elements)")


def add_qr_sidebar(tree, root):
    """
    Extend the SVG canvas to the right and add a dark QR sidebar.
    The original artwork is NOT moved — we just widen the viewBox/width.

    The sidebar rect starts 1 px to the LEFT of orig_w so it overlaps the
    ticket background by one pixel, which eliminates the hairline gap that
    appears in some renderers due to sub-pixel rounding.
    """
    svg_ns = _detect_namespace(root)
    tag = lambda name: f'{{{svg_ns}}}{name}' if svg_ns else name

    # --- Read current dimensions ---
    raw_w = root.get('width', '0')
    raw_h = root.get('height', '0')

    def to_float(val):
        try:
            return float(re.sub(r'[^0-9.]', '', val))
        except ValueError:
            return 0.0

    orig_w = to_float(raw_w)
    orig_h = to_float(raw_h)

    # Fall back to viewBox if width/height are missing or zero
    if orig_w == 0 or orig_h == 0:
        vb = root.get('viewBox', '')
        if vb:
            parts = vb.split()
            if len(parts) == 4:
                orig_w = float(parts[2])
                orig_h = float(parts[3])

    if orig_w == 0:
        orig_w = 1200
    if orig_h == 0:
        orig_h = 600

    print(f"\n📐 Original dimensions: {orig_w:.0f} × {orig_h:.0f}")

    # Clip any invisible out-of-bounds elements so the sidebar sits flush
    root.set('overflow', 'hidden')
    clip_original_content(root, orig_w, orig_h)

    # --- Sidebar geometry ---
    sidebar_w = max(360, orig_h * 0.30)   # proportional to ticket height
    new_w     = orig_w + sidebar_w
    pad       = sidebar_w * 0.10          # inner padding

    qr_size   = sidebar_w - pad * 2       # QR square fills most of the sidebar
    qr_x      = orig_w + pad
    qr_y      = (orig_h - qr_size) / 2   # vertically centred

    # --- Update root dimensions ---
    root.set('width',   f'{new_w:.0f}')
    root.set('height',  f'{orig_h:.0f}')
    root.set('viewBox', f'0 0 {new_w:.0f} {orig_h:.0f}')

    # --- Build sidebar group ---
    g = ET.SubElement(root, tag('g'))
    g.set('id', 'qr_sidebar')

    def sub(parent, name, attribs, text_content=None):
        el = ET.SubElement(parent, tag(name), attribs)
        if text_content is not None:
            el.text = text_content
        return el

    # Sidebar background — starts 1 px left of orig_w to close the hairline gap.
    # The extra pixel is hidden under the ticket's own background.
    OVERLAP = 1
    sub(g, 'rect', {
        'x':      str(orig_w - OVERLAP),
        'y':      '0',
        'width':  str(sidebar_w + OVERLAP),
        'height': str(orig_h),
        'fill':   '#000000',
    })

    # White background square for the QR code
    sub(g, 'rect', {
        'x':      str(qr_x),
        'y':      str(qr_y),
        'width':  str(qr_size),
        'height': str(qr_size),
        'fill':   '#ffffff',
        'rx':     '8',
    })

    # QR code image (will be filled at send-time)
    img = ET.SubElement(g, tag('image'))
    inner_pad = qr_size * 0.04
    img.set('x',      str(qr_x + inner_pad))
    img.set('y',      str(qr_y + inner_pad))
    img.set('width',  str(qr_size - inner_pad * 2))
    img.set('height', str(qr_size - inner_pad * 2))
    img.set('href',   '{QR_CODE_DATA}')
    img.set('{http://www.w3.org/1999/xlink}href', '{QR_CODE_DATA}')
    img.set('preserveAspectRatio', 'xMidYMid meet')

    print(f"✓ Canvas extended to {new_w:.0f} × {orig_h:.0f}")
    print(f"✓ QR sidebar added (width={sidebar_w:.0f}, QR size={qr_size:.0f}×{qr_size:.0f})")
    print(f"✓ Sidebar overlaps ticket by {OVERLAP} px — hairline gap eliminated")

    return tree


# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------

def save_svg(tree, output_path):
    try:
        ET.register_namespace('', 'http://www.w3.org/2000/svg')
        ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')
        tree.write(output_path, encoding='unicode', xml_declaration=False)
        print(f"\n✓ Saved → {output_path}")
        return True
    except Exception as e:
        print(f"\n❌ Error saving SVG: {e}")
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python prepare_svg.py your_ticket.svg")
        return

    input_file  = sys.argv[1]
    output_file = "ticket_template.svg"

    result = analyze_svg(input_file)
    if result[0] is None:
        return
    tree, root, text_elements = result

    print("\n" + "─" * 60)
    print("\n🔧 Setting up placeholders...\n")

    if text_elements:
        name_choice = input("Which element is the STUDENT NAME? (number, or 0 to skip): ").strip()
        name_index  = int(name_choice) if name_choice.isdigit() and int(name_choice) > 0 else None

        class_choice = input("Which element is the STUDENT CLASS? (number, or 0 to skip): ").strip()
        class_index  = int(class_choice) if class_choice.isdigit() and int(class_choice) > 0 else None

        if name_index or class_index:
            add_placeholders(text_elements, name_index, class_index)
    else:
        print("⚠️  Skipping placeholder step — no text elements found.")
        print("   Add {NAME_PLACEHOLDER} and {CLASS_PLACEHOLDER} manually to the SVG.\n")

    qr_choice = input("\nExtend canvas and add QR sidebar? (y/n): ").strip().lower()
    if qr_choice == 'y':
        tree = add_qr_sidebar(tree, root)

    print("\n" + "─" * 60)
    if save_svg(tree, output_file):
        print(f"""
✨ Done! Next steps:

1. Open {output_file} in a browser to check the layout
2. Confirm these placeholders are present:
     {{NAME_PLACEHOLDER}}   — student name
     {{CLASS_PLACEHOLDER}}  — student class
     {{QR_CODE_DATA}}       — QR code image (in the sidebar)
3. Run:  python main_svg.py

💡 If text is still missing, your Figma SVG may have text converted to
   outlines (paths). Re-export from Figma WITHOUT "Outline text".

💡 To re-run with different offsets, just re-run this script on the
   original file — don't re-run on the already-prepared template.
""")


if __name__ == "__main__":
    main()