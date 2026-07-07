"""Shared SVG → OOXML engine for PaperJSX v1.3 (PPT-P0-1).

Provides a minimal but usable SVG-to-DrawingML converter used by both PPT
(free_shape, background_shape) and DOCX (svg_shape).

Supported SVG features:
- Path commands: M/m L/l H/h V/v Q/q T/t C/c S/s Z
- Basic shapes → path: <rect>, <circle>, <ellipse>, <polygon>, <polyline>, <line>
- fill: solid, none; linearGradient (simple linear); radialGradient (basic)
- stroke: width, color, none (mapped to a:ln)
- opacity: fill-opacity, stroke-opacity (best-effort via alpha)
- transform: translate, scale (subset sufficient for typical decorative SVGs)

Not supported (by design): clipPath, mask, complex filters, text-on-path,
animate, foreignObject. Gradients use best-effort conversion to a:gradFill.

@since v1.3.0 (ARCH-1 / PPT-P0-1)
"""
from __future__ import annotations

import math
import re
from typing import Optional, Tuple
from xml.etree import ElementTree as ET

from lxml import etree


# ---------------------------------------------------------------------------
# XML namespaces
# ---------------------------------------------------------------------------
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS_DECL = "{%s}" % A_NS
P_NS_DECL = "{%s}" % P_NS
SVG_NS = "http://www.w3.org/2000/svg"

nsmap = {"a": A_NS, "p": P_NS, "r": R_NS}


# ---------------------------------------------------------------------------
# Minimal SVG parser: produces a normalized list of (path_data, style)
# ---------------------------------------------------------------------------

_NUMBER_RE = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
_PATH_CMD_RE = re.compile(r"([MmLlHhVvQqTtCcSsZzAa])|(-?\d*\.?\d+(?:[eE][-+]?\d+)?)")


def _parse_length(v: str, *, relative_to: float = 1.0, dpi: float = 96.0) -> float:
    """Convert an SVG length (e.g. "12", "12px", "50%", "2cm") to user units.

    Relative units ("%") are scaled by ``relative_to``.
    Absolute units assume 96 DPI.
    """
    if v is None:
        return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    if s.endswith("%"):
        return float(s[:-1]) / 100.0 * relative_to
    if s.endswith("px"):
        return float(s[:-2])
    if s.endswith("pt"):
        return float(s[:-2]) * dpi / 72.0
    if s.endswith("pc"):
        return float(s[:-2]) * dpi / 6.0
    if s.endswith("mm"):
        return float(s[:-2]) * dpi / 25.4
    if s.endswith("cm"):
        return float(s[:-2]) * dpi / 2.54
    if s.endswith("in"):
        return float(s[:-2]) * dpi
    return float(s)


def _parse_color(c: str) -> Optional[Tuple[int, int, int, float]]:
    """Parse SVG color → (r,g,b,a 0..1). Returns None for 'none'."""
    if c is None:
        return None
    c = c.strip().lower()
    if c in ("none", "transparent", ""):
        return None
    # #RGB / #RRGGBB
    m = re.match(r"^#([0-9a-f]{3,8})$", c)
    if m:
        h = m.group(1)
        if len(h) == 3:
            r, g, b = int(h[0]*2, 16), int(h[1]*2, 16), int(h[2]*2, 16)
            a = 1.0
        elif len(h) == 4:
            r, g, b = int(h[0]*2,16), int(h[1]*2,16), int(h[2]*2,16)
            a = int(h[3]*2,16)/255.0
        elif len(h) == 6:
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            a = 1.0
        elif len(h) == 8:
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            a = int(h[6:8],16)/255.0
        else:
            return None
        return (r, g, b, a)
    # rgb(..) / rgba(..)
    m = re.match(r"^rgba?\(([^)]+)\)$", c)
    if m:
        parts = [p.strip() for p in m.group(1).split(",")]
        if len(parts) >= 3:
            def chan(v):
                v = v.strip()
                if v.endswith("%"):
                    return int(float(v[:-1])*255/100)
                return int(float(v))
            r = chan(parts[0]); g = chan(parts[1]); b = chan(parts[2])
            a = float(parts[3]) if len(parts) > 3 else 1.0
            return (r, g, b, a)
    # Named colors (small common set — fall back to black)
    NAMED = {
        "black":(0,0,0),"white":(255,255,255),"red":(255,0,0),"green":(0,128,0),
        "blue":(0,0,255),"yellow":(255,255,0),"gray":(128,128,128),"grey":(128,128,128),
        "orange":(255,165,0),"purple":(128,0,128),"navy":(0,0,128),"silver":(192,192,192),
        "transparent":None,
    }
    if c in NAMED:
        v = NAMED[c]
        return None if v is None else (v[0], v[1], v[2], 1.0)
    return (0, 0, 0, 1.0)


# ----------------------- path parsing --------------------------------------

def _tokenize_path(d: str) -> list:
    """Tokenize an SVG path 'd' string into (cmd, [coords...]) chunks."""
    tokens = []
    cmd = None
    coords: list[float] = []
    for m in _PATH_CMD_RE.finditer(d):
        letter, num = m.group(1), m.group(2)
        if letter is not None:
            if cmd is not None:
                tokens.append((cmd, coords))
            cmd = letter
            coords = []
        elif num is not None:
            coords.append(float(num))
    if cmd is not None:
        tokens.append((cmd, coords))
    return tokens


def _svg_path_to_absolute_commands(d: str) -> list:
    """Return list of absolute-coord commands (uppercase letter + flat coords).

    Handles conversion of relative → absolute, H/V → L, T/S → Q/C, normalizes
    arc-free path to M/L/Q/C/Z.
    """
    tokens = _tokenize_path(d)
    out = []
    cx, cy = 0.0, 0.0  # current point (absolute)
    sx, sy = 0.0, 0.0  # subpath start
    prev_q_ctrl = None  # for T shorthand
    prev_c_ctrl_end = None  # for S shorthand reflection
    prev_cmd = None

    for cmd, coords in tokens:
        is_rel = cmd.islower()
        C = cmd.upper()
        i = 0
        if C == "M":
            while i < len(coords):
                x, y = coords[i], coords[i+1]
                if is_rel:
                    x += cx; y += cy
                out.append(("M", x, y))
                cx = cy = sx = sy = 0
                cx, cy = x, y
                sx, sy = x, y
                i += 2
                # Subsequent pairs are treated as L (per SVG spec)
                C = "L"; is_rel = cmd.islower()
        elif C == "L":
            while i < len(coords):
                x, y = coords[i], coords[i+1]
                if is_rel: x += cx; y += cy
                out.append(("L", x, y))
                cx, cy = x, y; i += 2
        elif C == "H":
            while i < len(coords):
                x = coords[i]
                if is_rel: x += cx
                out.append(("L", x, cy))
                cx = x; i += 1
        elif C == "V":
            while i < len(coords):
                y = coords[i]
                if is_rel: y += cy
                out.append(("L", cx, y))
                cy = y; i += 1
        elif C == "C":
            while i+5 < len(coords):
                x1,y1,x2,y2,x,y = coords[i:i+6]
                if is_rel:
                    x1+=cx;y1+=cy; x2+=cx;y2+=cy; x+=cx;y+=cy
                out.append(("C", x1,y1,x2,y2,x,y))
                prev_c_ctrl_end = (x2, y2)
                cx, cy = x, y; i += 6
        elif C == "S":
            while i+3 < len(coords):
                x2,y2,x,y = coords[i:i+4]
                if is_rel:
                    x2+=cx; y2+=cy; x+=cx; y+=cy
                # Reflect previous C's second control, or use current point
                if prev_cmd == "C" and prev_c_ctrl_end is not None:
                    x1 = 2*cx - prev_c_ctrl_end[0]
                    y1 = 2*cy - prev_c_ctrl_end[1]
                else:
                    x1, y1 = cx, cy
                out.append(("C", x1,y1,x2,y2,x,y))
                prev_c_ctrl_end = (x2, y2)
                cx, cy = x, y; i += 4
        elif C == "Q":
            while i+3 < len(coords):
                x1,y1,x,y = coords[i:i+4]
                if is_rel:
                    x1+=cx;y1+=cy;x+=cx;y+=cy
                out.append(("Q", x1,y1,x,y))
                prev_q_ctrl = (x1, y1)
                cx, cy = x, y; i += 4
        elif C == "T":
            while i+1 < len(coords):
                x, y = coords[i], coords[i+1]
                if is_rel:
                    x += cx; y += cy
                if prev_cmd == "Q" and prev_q_ctrl is not None:
                    x1 = 2*cx - prev_q_ctrl[0]
                    y1 = 2*cy - prev_q_ctrl[1]
                else:
                    x1, y1 = cx, cy
                out.append(("Q", x1,y1,x,y))
                prev_q_ctrl = (x1, y1)
                cx, cy = x, y; i += 2
        elif C == "Z":
            out.append(("Z",))
            cx, cy = sx, sy
        elif C == "A":
            # SVG arc → approximate with cubic Bezier (cheap; acceptable for decor)
            while i+6 < len(coords):
                rx, ry, rot, large_arc, sweep, x, y = coords[i:i+7]
                large_arc = bool(large_arc); sweep = bool(sweep)
                if is_rel: x += cx; y += cy
                _arc_to_cubic(out, cx, cy, rx, ry, rot, large_arc, sweep, x, y)
                cx, cy = x, y; i += 7
        # reset state for non-curve commands
        if C not in ("C", "S"):
            prev_c_ctrl_end = None
        if C not in ("Q", "T"):
            prev_q_ctrl = None
        prev_cmd = C
    return out


def _arc_to_cubic(out, x1, y1, rx, ry, phi_deg, large_arc, sweep, x2, y2):
    """Approximate an SVG arc with cubic Bezier segments and append to out."""
    # Simple implementation based on https://mortoray.com/2017/02/16/rendering-an-svg-elliptical-arc-as-bezier-curves/
    if rx == 0 or ry == 0:
        out.append(("L", x2, y2))
        return
    phi = math.radians(phi_deg or 0.0)
    # Compute center
    dx = (x1 - x2) / 2.0
    dy = (y1 - y2) / 2.0
    x1p = dx * math.cos(phi) + dy * math.sin(phi)
    y1p = -dx * math.sin(phi) + dy * math.cos(phi)
    rxs = rx*rx; rys = ry*ry
    x1ps = x1p*x1p; y1ps = y1p*y1p
    lam = x1ps/rxs + y1ps/rys
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s; ry *= s; rxs = rx*rx; rys = ry*ry
    sign = -1 if large_arc == sweep else 1
    num = max(0.0, rxs*rys - rxs*y1ps - rys*x1ps)
    den = rxs*y1ps + rys*x1ps
    coef = sign * math.sqrt(num / den) if den > 0 else 0
    cxp = coef *  rx*y1p/ry
    cyp = coef * -ry*x1p/rx
    cx = cxp*math.cos(phi) - cyp*math.sin(phi) + (x1+x2)/2.0
    cy = cxp*math.sin(phi) + cyp*math.cos(phi) + (y1+y2)/2.0
    # Angles
    def angle(ux, uy, vx, vy):
        dot = ux*vx + uy*vy
        len_u = math.hypot(ux, uy); len_v = math.hypot(vx, vy)
        if len_u == 0 or len_v == 0:
            return 0.0
        a = math.acos(max(-1.0, min(1.0, dot/(len_u*len_v))))
        if ux*vy - uy*vx < 0:
            a = -a
        return a
    theta1 = angle(1, 0, (x1p-cxp)/rx, (y1p-cyp)/ry)
    dtheta = angle((x1p-cxp)/rx, (y1p-cyp)/ry, (-x1p-cxp)/rx, (-y1p-cyp)/ry)
    if not sweep and dtheta > 0:
        dtheta -= 2*math.pi
    elif sweep and dtheta < 0:
        dtheta += 2*math.pi
    # Break into segments
    n_segs = int(math.ceil(abs(dtheta) / (math.pi/2)))
    dtheta_seg = dtheta / n_segs
    t = 8/3 * math.sin(dtheta_seg/4)**2 / math.sin(dtheta_seg/2)
    x = x1; y = y1
    theta = theta1
    cos_phi = math.cos(phi); sin_phi = math.sin(phi)
    for _ in range(n_segs):
        cos1 = math.cos(theta); sin1 = math.sin(theta)
        cos2 = math.cos(theta + dtheta_seg); sin2 = math.sin(theta + dtheta_seg)
        # Endpoint of segment
        e_x = cx + rx*cos2*cos_phi - ry*sin2*sin_phi
        e_y = cy + rx*cos2*sin_phi + ry*sin2*cos_phi
        # Control points relative to arc endpoints
        c1x = x + t*(-rx*sin1*cos_phi - ry*cos1*sin_phi)
        c1y = y + t*(-rx*sin1*sin_phi + ry*cos1*cos_phi)
        c2x = e_x + t*(rx*sin2*cos_phi + ry*cos2*sin_phi)
        c2y = e_y + t*(rx*sin2*sin_phi - ry*cos2*cos_phi)
        out.append(("C", c1x, c1y, c2x, c2y, e_x, e_y))
        x, y = e_x, e_y
        theta += dtheta_seg


# ----------------------- shape → path helpers ------------------------------

def _rect_to_path(x, y, w, h, rx, ry):
    if rx == 0 and ry == 0:
        return f"M{x} {y} L{x+w} {y} L{x+w} {y+h} L{x} {y+h} Z"
    # rounded rect with arc corners
    rx = min(rx, w/2); ry = min(ry, h/2)
    return (
        f"M{x+rx} {y} "
        f"L{x+w-rx} {y} "
        f"A{rx} {ry} 0 0 1 {x+w} {y+ry} "
        f"L{x+w} {y+h-ry} "
        f"A{rx} {ry} 0 0 1 {x+w-rx} {y+h} "
        f"L{x+rx} {y+h} "
        f"A{rx} {ry} 0 0 1 {x} {y+h-ry} "
        f"L{x} {y+ry} "
        f"A{rx} {ry} 0 0 1 {x+rx} {y} Z"
    )


def _ellipse_to_path(cx, cy, rx, ry):
    # Two arc approach
    return (
        f"M{cx-rx} {cy} "
        f"A{rx} {ry} 0 1 0 {cx+rx} {cy} "
        f"A{rx} {ry} 0 1 0 {cx-rx} {cy} Z"
    )


def _circle_to_path(cx, cy, r):
    return _ellipse_to_path(cx, cy, r, r)


def _polyline_to_path(points_str, close=False):
    pts = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", points_str or "")
    if len(pts) < 4:
        return ""
    nums = [float(p) for p in pts]
    coords = [(nums[i], nums[i+1]) for i in range(0, len(nums)-1, 2)]
    if not coords:
        return ""
    out = "M{} {}".format(coords[0][0], coords[0][1])
    for x, y in coords[1:]:
        out += " L{} {}".format(x, y)
    if close:
        out += " Z"
    return out


def _line_to_path(x1, y1, x2, y2):
    return f"M{x1} {y1} L{x2} {y2}"


# ----------------------- Style resolution ----------------------------------

def _inherit_style(el: ET.Element, props: dict) -> dict:
    """Collect fill/stroke/fill-opacity/stroke-opacity/stroke-width, honoring
    presentation attributes and inline style="" attribute."""
    style = dict(props)
    # inline style overrides
    s = el.get("style")
    if s:
        for kv in s.split(";"):
            if ":" in kv:
                k, v = kv.split(":", 1)
                k = k.strip(); v = v.strip()
                # translate SVG style keys to our props dict keys
                kmap = {"fill":"fill","stroke":"stroke","stroke-width":"stroke_width",
                        "fill-opacity":"fill_opacity","stroke-opacity":"stroke_opacity",
                        "opacity":"opacity", "stop-color":"stop_color",
                        "stop-opacity":"stop_opacity"}
                if k in kmap:
                    style[kmap[k]] = v
    for svg_k, key in (("fill","fill"),("stroke","stroke"),
                        ("stroke-width","stroke_width"),
                        ("fill-opacity","fill_opacity"),
                        ("stroke-opacity","stroke_opacity"),
                        ("opacity","opacity")):
        v = el.get(svg_k)
        if v is not None:
            style[key] = v
    return style


# ----------------------- Gradient parsing ----------------------------------

def _parse_gradient(defs_el: ET.Element, ref: str, bbox: Tuple[float,float,float,float]) -> Optional[dict]:
    if not ref.startswith("url("):
        return None
    m = re.match(r"url\(#([^)]+)\)", ref.strip())
    if not m:
        return None
    gid = m.group(1)
    # find by id
    for el in defs_el.iter() if defs_el is not None else []:
        if el.get("id") == gid:
            return _build_gradient(el, bbox)
    # search globally in the root (fallback)
    return None


def _build_gradient(el: ET.Element, bbox: Tuple[float,float,float,float]) -> dict:
    x, y, w, h = bbox
    tag = el.tag.split("}")[-1]
    stops = []
    for st in el.findall(".//{http://www.w3.org/2000/svg}stop"):
        off = st.get("offset", "0")
        try:
            if off.endswith("%"):
                op = float(off[:-1])/100.0
            else:
                op = float(off)
        except Exception:
            op = 0.0
        col = _parse_color(st.get("stop-color", "#000"))
        opa = st.get("stop-opacity")
        if col is None:
            col = (0,0,0,1)
        a = col[3]
        if opa is not None:
            try: a = float(opa)
            except: pass
        stops.append({"offset": op, "color": (col[0], col[1], col[2], a)})
    if not stops:
        stops = [{"offset":0,"color":(0,0,0,1)}, {"offset":1,"color":(255,255,255,1)}]
    if tag == "linearGradient":
        x1 = _parse_length(el.get("x1","0%"), relative_to=w) + x
        y1 = _parse_length(el.get("y1","0%"), relative_to=h) + y
        x2 = _parse_length(el.get("x2","100%"), relative_to=w) + x
        y2 = _parse_length(el.get("y2","0%"), relative_to=h) + y
        return {"type":"linear", "x1":x1,"y1":y1,"x2":x2,"y2":y2, "stops":stops}
    if tag == "radialGradient":
        cx = _parse_length(el.get("cx","50%"), relative_to=w) + x
        cy = _parse_length(el.get("cy","50%"), relative_to=h) + y
        r = _parse_length(el.get("r","50%"), relative_to=min(w,h)/2 if min(w,h)>0 else 1)
        fx = _parse_length(el.get("fx", str(cx)), relative_to=w) + x
        fy = _parse_length(el.get("fy", str(cy)), relative_to=h) + y
        return {"type":"radial", "cx":cx,"cy":cy,"r":r,"fx":fx,"fy":fy, "stops":stops}
    return None


# ----------------------- Top-level SVG parse -------------------------------

def _collect_shapes(root: ET.Element, defs: Optional[ET.Element], parent_style: dict,
                    transforms: list, shapes: list):
    """Recursively walk SVG tree, collecting shape entries with path+style."""
    for child in list(root):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("defs", "metadata", "title", "desc", "style"):
            # Merge defs into passed-defs
            continue
        style = _inherit_style(child, parent_style)
        # transform
        tr = child.get("transform")
        new_tf = list(transforms)
        if tr:
            new_tf.append(tr)
        d = None
        if tag == "path":
            d = child.get("d")
        elif tag == "rect":
            x = _parse_length(child.get("x","0")); y = _parse_length(child.get("y","0"))
            w = _parse_length(child.get("width","0")); h = _parse_length(child.get("height","0"))
            rx = _parse_length(child.get("rx",child.get("ry","0"))); ry = _parse_length(child.get("ry",child.get("rx","0")))
            d = _rect_to_path(x, y, w, h, rx, ry)
        elif tag == "circle":
            cx = _parse_length(child.get("cx","0")); cy = _parse_length(child.get("cy","0"))
            r = _parse_length(child.get("r","0"))
            d = _circle_to_path(cx, cy, r)
        elif tag == "ellipse":
            cx = _parse_length(child.get("cx","0")); cy = _parse_length(child.get("cy","0"))
            rx = _parse_length(child.get("rx","0")); ry = _parse_length(child.get("ry","0"))
            d = _ellipse_to_path(cx, cy, rx, ry)
        elif tag == "polygon":
            d = _polyline_to_path(child.get("points",""), close=True)
        elif tag == "polyline":
            d = _polyline_to_path(child.get("points",""), close=False)
        elif tag == "line":
            x1 = _parse_length(child.get("x1","0")); y1 = _parse_length(child.get("y1","0"))
            x2 = _parse_length(child.get("x2","0")); y2 = _parse_length(child.get("y2","0"))
            d = _line_to_path(x1,y1,x2,y2)
        elif tag == "g":
            _collect_shapes(child, defs, style, new_tf, shapes)
            continue
        else:
            # skip unknown (text/image etc.)
            continue
        if d:
            shapes.append({"d": d, "style": style, "transforms": new_tf})
        _collect_shapes(child, defs, style, new_tf, shapes)


def _apply_transforms(cmds: list, transforms: list) -> list:
    """Apply SVG transform="translate(...) scale(...)" list to absolute commands."""
    # Build affine matrix (a b c d e f)
    a, b, c, dd, e, f = 1.0, 0.0, 0.0, 1.0, 0.0, 0.0
    for tr in transforms:
        # parse translate/scale/rotate/matrix (subset)
        for name, args in re.findall(r"(translate|scale|rotate|matrix|skewX|skewY)\s*\(([^)]*)\)", tr):
            nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", args)]
            if name == "translate":
                tx, ty = nums[0], nums[1] if len(nums) > 1 else 0.0
                a,b,c,dd,e,f = a,b,c,dd,e+tx,f+ty
            elif name == "scale":
                sx, sy = (nums[0], nums[1]) if len(nums)>1 else (nums[0], nums[0])
                a,b,c,dd,e,f = a*sx, b*sx, c*sy, dd*sy, e, f
            elif name == "matrix" and len(nums) == 6:
                a1,b1,c1,d1,e1,f1 = nums
                a,b,c,dd,e,f = a*a1+c*b1, b*a1+dd*b1, a*c1+c*d1, b*c1+dd*d1, a*e1+c*f1+e, b*e1+dd*f1+f
            elif name == "rotate":
                ang = math.radians(nums[0])
                cos_a = math.cos(ang); sin_a = math.sin(ang)
                if len(nums) == 3:
                    cxr, cyr = nums[1], nums[2]
                    # translate(-cx,-cy) rotate translate(cx,cy)
                    e -= cxr; f -= cyr
                    a2,b2,c2,d2,e2,f2 = cos_a, sin_a, -sin_a, cos_a, 0, 0
                    a,b,c,dd,e,f = a*a2+c*b2, b*a2+dd*b2, a*c2+c*d2, b*c2+dd*d2, a*e2+c*f2+e, b*e2+dd*f2+f
                    e += cxr; f += cyr
                else:
                    a2,b2,c2,d2 = cos_a, sin_a, -sin_a, cos_a
                    a,b,c,dd,e,f = a*a2+c*b2, b*a2+dd*b2, a*c2+c*d2, b*c2+dd*d2, e, f
            # skewX/skewY best-effort
            elif name == "skewX" and nums:
                t = math.tan(math.radians(nums[0]))
                a,b,c,dd,e,f = a+c*t, b+dd*t, c, dd, e, f
            elif name == "skewY" and nums:
                t = math.tan(math.radians(nums[0]))
                a,b,c,dd,e,f = a, b, c+a*t, dd+b*t, e, f
    out = []
    for cmd in cmds:
        code = cmd[0]
        if code == "M":
            x,y = cmd[1], cmd[2]
            out.append(("M", a*x+c*y+e, b*x+dd*y+f))
        elif code == "L":
            x,y = cmd[1], cmd[2]
            out.append(("L", a*x+c*y+e, b*x+dd*y+f))
        elif code == "Q":
            x1,y1,x,y = cmd[1],cmd[2],cmd[3],cmd[4]
            out.append(("Q", a*x1+c*y1+e, b*x1+dd*y1+f, a*x+c*y+e, b*x+dd*y+f))
        elif code == "C":
            x1,y1,x2,y2,x,y = cmd[1],cmd[2],cmd[3],cmd[4],cmd[5],cmd[6]
            out.append(("C",
                a*x1+c*y1+e, b*x1+dd*y1+f,
                a*x2+c*y2+e, b*x2+dd*y2+f,
                a*x+c*y+e, b*x+dd*y+f))
        elif code == "Z":
            out.append(("Z",))
    return out


def _bbox_of_commands(all_cmds: list) -> Tuple[float,float,float,float]:
    xs=[]; ys=[]
    for cmds in all_cmds:
        for cmd in cmds:
            c = cmd[0]
            if c == "M" or c == "L":
                xs.append(cmd[1]); ys.append(cmd[2])
            elif c == "Q":
                xs.extend([cmd[1], cmd[3]]); ys.extend([cmd[2], cmd[4]])
            elif c == "C":
                xs.extend([cmd[1], cmd[3], cmd[5]]); ys.extend([cmd[2], cmd[4], cmd[6]])
    if not xs:
        return (0,0,1,1)
    return (min(xs), min(ys), max(xs)-min(xs), max(ys)-min(ys))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def svg_to_ooxml(svg_str: str, target_width_emu: int, target_height_emu: int,
                 *, shape_id: int = 1, name: str = "FreeShape",
                 x_emu: int = 0, y_emu: int = 0,
                 default_fill: str = "#2E75B6",
                 default_stroke: str = "none") -> etree._Element:
    """Convert an SVG string to a ``p:sp`` element that can be appended to a
    slide shape tree.

    Parameters
    ----------
    svg_str : str
        Full <svg ...>...</svg> content.
    target_width_emu / target_height_emu : int
        Target EMU width/height. Coordinates are scaled to fit preserving aspect
        unless the SVG has no viewBox (uses w/h attributes directly).
    x_emu / y_emu : int
        Position (EMU) of the shape in the slide.
    shape_id, name : int, str
        OOXML shape id/name.

    Returns
    -------
    lxml.etree._Element
        A ``p:sp`` element (with ``p:spPr/a:custGeom``) ready to be appended
        into ``slide.shapes._spTree``.
    """
    # Strip default namespaces if user supplied without (ET needs the {NS} form)
    txt = svg_str.strip()
    if not txt.startswith("<"):
        # treat as file path (best effort)
        try:
            with open(txt, "r", encoding="utf-8") as f:
                txt = f.read()
        except Exception:
            pass
    # Ensure SVG namespace for parsing
    if "xmlns=" not in txt[:200]:
        txt = txt.replace("<svg", "<svg xmlns=\"http://www.w3.org/2000/svg\"", 1)
    try:
        root = ET.fromstring(txt.encode("utf-8"))
    except ET.ParseError as e:
        raise ValueError(f"Invalid SVG: {e}")
    vb = root.get("viewBox")
    svg_w = _parse_length(root.get("width", "100%"), relative_to=100)
    svg_h = _parse_length(root.get("height", "100%"), relative_to=100)
    if vb:
        parts = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", vb)]
        if len(parts) == 4:
            ox, oy, vw, vh = parts
        else:
            ox, oy, vw, vh = 0, 0, svg_w, svg_h
    else:
        ox, oy, vw, vh = 0, 0, svg_w or 100, svg_h or 100

    defs = root.find("{http://www.w3.org/2000/svg}defs")
    default_style = {"fill": default_fill, "stroke": default_stroke,
                     "stroke_width": "1", "fill_opacity": "1",
                     "stroke_opacity": "1", "opacity": "1"}
    shapes = []
    _collect_shapes(root, defs, default_style, [], shapes)

    # Parse each shape into abs commands and apply transforms
    shape_cmds = []
    shape_styles = []
    for s in shapes:
        cmds = _svg_path_to_absolute_commands(s["d"])
        cmds = _apply_transforms(cmds, s["transforms"])
        shape_cmds.append(cmds)
        shape_styles.append(s["style"])

    if not shape_cmds:
        # empty — create a no-op placeholder
        shape_cmds = [[("M", 0,0), ("L", vw,0), ("L", vw,vh), ("L",0,vh), ("Z",)]]
        shape_styles = [dict(default_style, fill="none")]

    # Compute bbox in user coords for gradient/normalization
    bx, by, bw, bh = _bbox_of_commands(shape_cmds)
    # Determine path w/h in EMU local units: target
    if bw == 0: bw = 1
    if bh == 0: bh = 1
    sx = target_width_emu / bw
    sy = target_height_emu / bh
    scl = min(sx, sy)
    path_w = int(bw * scl)
    path_h = int(bh * scl)

    # Build p:sp element
    sp = etree.SubElement(etree.Element("dummy"), f"{{{P_NS}}}sp", nsmap=nsmap)
    sp = etree.Element(f"{{{P_NS}}}sp", nsmap=nsmap)

    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr", id=str(shape_id), name=name)
    etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")
    etree.SubElement(nvSpPr, f"{{{P_NS}}}nvPr")

    spPr = etree.SubElement(sp, f"{{{P_NS}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    etree.SubElement(xfrm, f"{{{A_NS}}}off", x=str(x_emu), y=str(y_emu))
    etree.SubElement(xfrm, f"{{{A_NS}}}ext", cx=str(target_width_emu), cy=str(target_height_emu))

    custGeom = etree.SubElement(spPr, f"{{{A_NS}}}custGeom")
    etree.SubElement(custGeom, f"{{{A_NS}}}avLst")
    etree.SubElement(custGeom, f"{{{A_NS}}}gdLst")
    etree.SubElement(custGeom, f"{{{A_NS}}}ahLst")
    etree.SubElement(custGeom, f"{{{A_NS}}}cxnLst")
    etree.SubElement(custGeom, f"{{{A_NS}}}rect", l="l", t="t", r="r", b="b")
    pathLst = etree.SubElement(custGeom, f"{{{A_NS}}}pathLst")
    pathEl = etree.SubElement(pathLst, f"{{{A_NS}}}path", w=str(path_w), h=str(path_h))

    # Helper to convert user coord to EMU in path-local space
    def to_emu_x(u: float) -> int:
        return int((u - bx) * scl)
    def to_emu_y(u: float) -> int:
        return int((u - by) * scl)

    # Determine fill/stroke overall: we create a single path per shape inside
    # its own <a:path>? No — DrawingML custGeom allows multiple moveTo segments
    # but fill/stroke is applied at spPr level for the whole shape.  To preserve
    # per-path style, we group shapes by fill+stroke and create spPr per path
    # segment is complex; instead, create MULTIPLE <p:sp> siblings — but API
    # returns a single element. We return a <p:grpSp> group when there are
    # multiple sub-shapes with different styles.
    def _append_path_cmds(parent_path, cmds):
        for cmd in cmds:
            c = cmd[0]
            if c == "M":
                m = etree.SubElement(parent_path, f"{{{A_NS}}}moveTo")
                etree.SubElement(m, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[1])), y=str(to_emu_y(cmd[2])))
            elif c == "L":
                ln = etree.SubElement(parent_path, f"{{{A_NS}}}lnTo")
                etree.SubElement(ln, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[1])), y=str(to_emu_y(cmd[2])))
            elif c == "Q":
                q = etree.SubElement(parent_path, f"{{{A_NS}}}quadBezTo")
                etree.SubElement(q, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[1])), y=str(to_emu_y(cmd[2])))
                etree.SubElement(q, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[3])), y=str(to_emu_y(cmd[4])))
            elif c == "C":
                cb = etree.SubElement(parent_path, f"{{{A_NS}}}cubicBezTo")
                etree.SubElement(cb, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[1])), y=str(to_emu_y(cmd[2])))
                etree.SubElement(cb, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[3])), y=str(to_emu_y(cmd[4])))
                etree.SubElement(cb, f"{{{A_NS}}}pt", x=str(to_emu_x(cmd[5])), y=str(to_emu_y(cmd[6])))
            elif c == "Z":
                etree.SubElement(parent_path, f"{{{A_NS}}}close")

    # For simplicity, render the FIRST shape's style at spPr level (common case:
    # one figure with fill/stroke) and append all path commands. Multi-style
    # SVGs render with the first shape's fill/stroke. (Advanced multi-fill
    # requires group shapes — TODO if needed.)
    main_style = shape_styles[0] if shape_styles else default_style
    all_cmds_merged = []
    for cmds in shape_cmds:
        all_cmds_merged.extend(cmds)
    # Split at M boundaries into path segments within one <a:path>
    for seg in _split_at_move(all_cmds_merged):
        _append_path_cmds(pathEl, seg)

    # Apply fill
    fill = main_style.get("fill", default_fill)
    _apply_fill(spPr, fill, defs, (bx, by, bw, bh), main_style)

    # Apply stroke
    stroke = main_style.get("stroke", default_stroke)
    sw = _parse_length(main_style.get("stroke_width", "1"), relative_to=1)
    _apply_stroke(spPr, stroke, sw, main_style)

    # effectLst (simple outer shadow, optional — off by default to avoid bloat)
    # We leave a:effectLst empty (no shadow) — renderer may add shadows on top.

    # p:style (required for consistent appearance)
    style = etree.SubElement(sp, f"{{{P_NS}}}style")
    lnRef = etree.SubElement(style, f"{{{A_NS}}}lnRef", idx="1")
    etree.SubElement(lnRef, f"{{{A_NS}}}schemeClr", val="accent1")
    fillRef = etree.SubElement(style, f"{{{A_NS}}}fillRef", idx="0")
    etree.SubElement(fillRef, f"{{{A_NS}}}schemeClr", val="accent1")
    effectRef = etree.SubElement(style, f"{{{A_NS}}}effectRef", idx="0")
    etree.SubElement(effectRef, f"{{{A_NS}}}schemeClr", val="accent1")
    fontRef = etree.SubElement(style, f"{{{A_NS}}}fontRef", idx="minor")
    etree.SubElement(fontRef, f"{{{A_NS}}}schemeClr", val="lt1")

    # Empty txBody (shapes are decorative; no text)
    txBody = etree.SubElement(sp, f"{{{P_NS}}}txBody")
    etree.SubElement(txBody, f"{{{A_NS}}}bodyPr", rtlCol="0", anchor="ctr")
    etree.SubElement(txBody, f"{{{A_NS}}}lstStyle")
    etree.SubElement(txBody, f"{{{A_NS}}}p")

    return sp


def _split_at_move(cmds: list) -> list:
    """Split command list at M boundaries, preserving Z/M continuation."""
    segs = []
    cur = []
    for cmd in cmds:
        if cmd[0] == "M" and cur:
            segs.append(cur)
            cur = [cmd]
        else:
            cur.append(cmd)
    if cur:
        segs.append(cur)
    return segs


def _apply_fill(spPr, fill_ref, defs, bbox, style):
    col = _parse_color(fill_ref)
    grad = _parse_gradient(defs, fill_ref, bbox) if fill_ref.startswith("url(") else None
    # overall opacity
    op_text = style.get("opacity", "1")
    try: op_factor = float(op_text)
    except: op_factor = 1.0
    if grad is not None:
        gradFill = etree.SubElement(spPr, f"{{{A_NS}}}gradFill", flip="none", rotWithShape="1")
        gsLst = etree.SubElement(gradFill, f"{{{A_NS}}}gsLst")
        for st in grad["stops"]:
            gs = etree.SubElement(gsLst, f"{{{A_NS}}}gs", pos=str(int(st["offset"]*100000)))
            r,g,b,a = st["color"]
            a = max(0, min(1, a*op_factor))
            srgb = etree.SubElement(gs, f"{{{A_NS}}}srgbClr", val=f"{r:02X}{g:02X}{b:02X}")
            etree.SubElement(srgb, f"{{{A_NS}}}alpha", val=str(int(a*100000)))
        if grad["type"] == "linear":
            x1,y1,x2,y2 = grad["x1"],grad["y1"],grad["x2"],grad["y2"]
            bx,by,bw,bh = bbox
            # Convert vector to angle (DrawingML lin uses ang attribute in 60000ths of degree)
            dx = x2-x1; dy = y2-y1
            ang = math.degrees(math.atan2(dy, dx))
            # DrawingML lin: ang is measured from horizontal left-to-right, going clockwise
            lin = etree.SubElement(gradFill, f"{{{A_NS}}}lin", ang=str(int(-ang*60000)), scaled="1")
        else:
            path = etree.SubElement(gradFill, f"{{{A_NS}}}path", path="circle")
            etree.SubElement(path, f"{{{A_NS}}}fillToRect", l="50000", t="50000", r="50000", b="50000")
        etree.SubElement(gradFill, f"{{{A_NS}}}tileRect")
    elif col is None:
        etree.SubElement(spPr, f"{{{A_NS}}}noFill")
    else:
        r,g,b,a = col
        a = max(0, min(1, a*op_factor))
        solid = etree.SubElement(spPr, f"{{{A_NS}}}solidFill")
        srgb = etree.SubElement(solid, f"{{{A_NS}}}srgbClr", val=f"{r:02X}{g:02X}{b:02X}")
        if a < 0.999:
            etree.SubElement(srgb, f"{{{A_NS}}}alpha", val=str(int(a*100000)))


def _apply_stroke(spPr, stroke_ref, width_pt, style):
    col = _parse_color(stroke_ref)
    if col is None:
        ln = etree.SubElement(spPr, f"{{{A_NS}}}ln", w=str(int(max(width_pt, 0)*12700)))
        etree.SubElement(ln, f"{{{A_NS}}}noFill")
        return
    w_emu = max(12700, int(width_pt * 12700))  # min 1pt
    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln", w=str(w_emu))
    r,g,b,a = col
    sop = style.get("stroke_opacity", "1")
    try: a = a*float(sop)
    except: pass
    op_text = style.get("opacity", "1")
    try: a = a*float(op_text)
    except: pass
    solid = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
    srgb = etree.SubElement(solid, f"{{{A_NS}}}srgbClr", val=f"{r:02X}{g:02X}{b:02X}")
    if a < 0.999:
        etree.SubElement(srgb, f"{{{A_NS}}}alpha", val=str(int(a*100000)))
    # round line ends/cap for a polished look
    ln.set("cap", "rnd")
    etree.SubElement(ln, f"{{{A_NS}}}round")


def svg_to_docx_drawing(svg_str: str, width_emu: int, height_emu: int,
                        *, shape_id: int = 1, name: str = "SvgShape",
                        doc_grid: int = 0) -> etree._Element:
    """Return a DOCX-compatible ``w:drawing`` element containing the SVG as a
    custom-geometry shape (using DrawingML, no image fallback).

    The element can be appended to a paragraph's run wrapper via
    ``run._r.append(drawing)``.
    """
    # Build a p:sp-like element but using wsp (wordprocessing shape)
    A = A_NS_DECL
    W = "{%s}" % W_NS
    WP = "{%s}" % WP_NS
    # We construct a standard wp:inline/wsp with a:custGeom identical to PPT
    nsmap = {"w": W_NS, "wp": WP_NS, "a": A_NS, "r": R_NS}
    drawing = etree.Element(f"{W}drawing", nsmap={k:v for k,v in nsmap.items()})
    inline = etree.SubElement(drawing, f"{WP}inline",
                              distT="0", distB="0", distL="0", distR="0")
    etree.SubElement(inline, f"{WP}extent", cx=str(width_emu), cy=str(height_emu))
    effectExtent = etree.SubElement(inline, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
    docPr = etree.SubElement(inline, f"{WP}docPr", id=str(shape_id), name=name)
    etree.SubElement(inline, f"{WP}cNvGraphicFramePr")
    graphic = etree.SubElement(inline, f"{A}graphic")
    graphicData = etree.SubElement(graphic, f"{A}graphicData",
                                   uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")

    # wsp: wordprocessing shape
    wsp = etree.SubElement(graphicData, f"{WP}wsp")
    nvSpPr = etree.SubElement(wsp, f"{WP}cNvSpPr")
    etree.SubElement(nvSpPr, f"{WP}cNvPr", id=str(shape_id+1000), name=name)
    etree.SubElement(nvSpPr, f"{WP}cNvSpPr")
    # spPr (re-using same custGeom)
    spPr = etree.SubElement(wsp, f"{A}spPr")
    xfrm = etree.SubElement(spPr, f"{A}xfrm")
    etree.SubElement(xfrm, f"{A}off", x="0", y="0")
    etree.SubElement(xfrm, f"{A}ext", cx=str(width_emu), cy=str(height_emu))

    # Reuse the PPT builder's custGeom/fill/stroke by temporarily rendering a
    # p:sp and stealing spPr contents.
    sp = svg_to_ooxml(svg_str, width_emu, height_emu, shape_id=shape_id, name=name)
    src_spPr = sp.find(f"{{{P_NS}}}spPr")
    # Copy children of src spPr (custGeom, fill, ln) into our DOCX spPr
    for child in src_spPr:
        spPr.append(etree.fromstring(etree.tostring(child)))

    # txBody (empty, same as PPT)
    txBody = etree.SubElement(wsp, f"{A}txBody")
    etree.SubElement(txBody, f"{A}bodyPr", rtlCol="0", anchor="ctr")
    etree.SubElement(txBody, f"{A}lstStyle")
    etree.SubElement(txBody, f"{A}p")

    return drawing


__all__ = [
    "svg_to_ooxml", "svg_to_docx_drawing",
    # v1.4 P2-1: SVG mini charts (sparklines) for KPI cards / table inline
    "mini_bar", "mini_line", "mini_pie", "mini_chart",
]


# ----------------------------------------------------------------------
# v1.4 P2-1: SVG 迷你图表（火花图）——纯 SVG 字符串，不依赖 matplotlib
# ----------------------------------------------------------------------

_MINI_CHART_PALETTE = [
    "primary", "accent", "secondary", "chart2", "chart3", "chart4",
]


def _resolve_color(color, theme_colors, idx: int = 0) -> str:
    """Resolve a color argument to a HEX (#rrggbb) string.

    - If ``color`` is an explicit string, return it (accepts #rgb/#rrggbb/named-hex).
    - Otherwise pick a color from ``theme_colors`` by cycling palette keys.
    """
    if color:
        c = str(color).strip()
        if c and not c.startswith("_"):
            return c if c.startswith("#") else f"#{c.lstrip('#')}"
    if not theme_colors:
        # fallback to academic primary
        return "#1F3864"
    # Allow passing a plain list of colors (convenience for mini_pie etc.)
    if isinstance(theme_colors, (list, tuple)):
        return theme_colors[idx % len(theme_colors)]
    # Try chart palette first (from themes.get_chart_palette) if provided
    chart_palette = theme_colors.get("_chart_palette")
    if chart_palette and isinstance(chart_palette, (list, tuple)):
        return chart_palette[idx % len(chart_palette)]
    key = _MINI_CHART_PALETTE[idx % len(_MINI_CHART_PALETTE)]
    c = theme_colors.get(key)
    if not c:
        # fall back: try primary/accent/secondary/text
        for k in ("primary", "accent", "secondary", "text"):
            if theme_colors.get(k):
                c = theme_colors[k]; break
    if not c:
        return "#1F3864"
    return c if c.startswith("#") else f"#{c.lstrip('#')}"


def _hesc(hex_color: str) -> str:
    """Ensure HEX is #rrggbb form."""
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(ch * 2 for ch in h)
    return f"#{h}"


def mini_bar(
    data,
    width: int = 120,
    height: int = 40,
    color=None,
    *,
    theme_colors=None,
    padding: int = 2,
) -> str:
    """Generate an SVG mini bar-chart (sparkline bars).

    Parameters
    ----------
    data : list[float]
        6~12 data points. Will be clamped to 3..12.
    width, height : int
        SVG viewport size in px.
    color : str, optional
        Explicit bar color (#rrggbb). If ``None``, cycle through theme_colors.
    theme_colors : dict, optional
        Theme palette dict from ``shared.themes.get_theme_palette`` (or a
        dict that also provides ``_chart_palette`` for multi-color).
    padding : int
        Inner padding in px.
    """
    vals = [float(v) for v in (data or [])]
    n = max(3, min(12, len(vals)))
    vals = vals[:n]
    vmin, vmax = min(vals), max(vals)
    if vmax == vmin:
        vmax = vmin + 1.0
    inner_w = width - padding * 2
    inner_h = height - padding * 2
    gap = max(1, int(inner_w * 0.12 / n))
    bar_w = max(2, (inner_w - gap * (n - 1)) / n)
    bars = []
    for i, v in enumerate(vals):
        bh = inner_h * (v - vmin) / (vmax - vmin)
        bh = max(1.5, bh)
        x = padding + i * (bar_w + gap)
        y = padding + inner_h - bh
        fill = _hesc(_resolve_color(color, theme_colors, idx=i))
        bars.append(
            f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_w:.2f}" height="{bh:.2f}" '
            f'rx="1" fill="{fill}" opacity="0.92"/>'
        )
    # baseline
    baseline_y = padding + inner_h
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" role="img" aria-label="mini bar chart">'
        f'<line x1="{padding}" y1="{baseline_y}" x2="{width-padding}" y2="{baseline_y}" '
        f'stroke="#cccccc" stroke-width="0.5"/>'
        f'{"".join(bars)}'
        f'</svg>'
    )


def mini_line(
    data,
    width: int = 120,
    height: int = 40,
    color=None,
    *,
    theme_colors=None,
    fill: bool = True,
    padding: int = 2,
    stroke_width: float = 1.5,
) -> str:
    """Generate an SVG mini line-chart (sparkline) with optional area fill."""
    vals = [float(v) for v in (data or [])]
    n = max(2, len(vals))
    vals = vals[:n]
    vmin, vmax = min(vals), max(vals)
    if vmax == vmin:
        vmax = vmin + 1.0
    inner_w = width - padding * 2
    inner_h = height - padding * 2
    step = inner_w / max(1, n - 1)
    pts = []
    for i, v in enumerate(vals):
        x = padding + i * step
        y = padding + inner_h - inner_h * (v - vmin) / (vmax - vmin)
        pts.append((x, y))
    line_color = _hesc(_resolve_color(color, theme_colors, idx=0))
    poly = " ".join(f"{x:.2f},{y:.2f}" for x, y in pts)
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}" role="img" aria-label="mini line chart">'
    ]
    # baseline
    baseline_y = padding + inner_h
    parts.append(
        f'<line x1="{padding}" y1="{baseline_y}" x2="{width-padding}" y2="{baseline_y}" '
        f'stroke="#cccccc" stroke-width="0.5"/>'
    )
    if fill:
        area = (
            f"M {pts[0][0]:.2f},{baseline_y:.2f} "
            + " ".join(f"L {x:.2f},{y:.2f}" for x, y in pts)
            + f" L {pts[-1][0]:.2f},{baseline_y:.2f} Z"
        )
        parts.append(f'<path d="{area}" fill="{line_color}" opacity="0.18"/>')
    parts.append(
        f'<polyline points="{poly}" fill="none" stroke="{line_color}" '
        f'stroke-width="{stroke_width}" stroke-linejoin="round" stroke-linecap="round"/>'
    )
    # last-point dot
    x, y = pts[-1]
    parts.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="1.6" fill="{line_color}"/>')
    parts.append("</svg>")
    return "".join(parts)


def mini_pie(
    parts,
    size: int = 80,
    colors=None,
    *,
    theme_colors=None,
    stroke: str = "#ffffff",
    stroke_width: float = 1.0,
) -> str:
    """Generate an SVG mini pie-chart using ``circle`` + ``stroke-dasharray`` trick.

    Parameters
    ----------
    parts : list[float] | dict
        Either a list of numeric values, or a dict whose values are numeric.
        Capped at 6 slices; zeros/negatives are skipped.
    size : int
        Diameter of the pie in px.
    colors : list[str], optional
        Explicit per-slice colors. If ``None``, cycle theme_colors.
    """
    if isinstance(parts, dict):
        items = [(k, float(v)) for k, v in parts.items()]
    else:
        items = [(None, float(v)) for v in parts or []]
    # filter non-positive and cap at 6 slices
    items = [(k, v) for k, v in items if v > 0][:6]
    total = sum(v for _, v in items)
    if total <= 0 or not items:
        # empty pie — draw a gray ring
        return (
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
            f'viewBox="0 0 {size} {size}" role="img" aria-label="mini pie chart (empty)">'
            f'<circle cx="{size/2}" cy="{size/2}" r="{size/2-1}" fill="#e5e7eb"/></svg>'
        )
    # circle-based pie: use a very wide stroke equal to r, circumference = 2πr
    cx = cy = size / 2
    r = size / 2 - 1  # radius of outer circle
    sw = r             # stroke width == r → half the diameter filled
    rr = r / 2         # path radius = r - sw/2 = r/2
    circum = 2 * math.pi * rr
    # stroke is drawn centered on a circle of radius rr, with stroke-width sw=r
    # so the resulting ring exactly covers the disk of radius r.
    segs = []
    offset = 0.0  # dashoffset moves starting point; SVG dash starts at 3 o'clock
    # rotate so the first slice starts at 12 o'clock (negative rotation)
    rotation = -90
    for i, (_, v) in enumerate(items):
        frac = v / total
        dash_len = frac * circum
        # explicit color lookup
        if colors and i < len(colors) and colors[i]:
            fill_c = _hesc(str(colors[i]))
        else:
            fill_c = _hesc(_resolve_color(None, theme_colors, idx=i))
        segs.append(
            f'<circle cx="{cx}" cy="{cy}" r="{rr}" fill="none" stroke="{fill_c}" '
            f'stroke-width="{sw}" '
            f'stroke-dasharray="{dash_len:.3f} {circum:.3f}" '
            f'stroke-dashoffset="{-offset:.3f}" '
            f'transform="rotate({rotation} {cx} {cy})"/>'
        )
        offset += dash_len
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 {size} {size}" role="img" aria-label="mini pie chart">'
        f'{"".join(segs)}'
        f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{stroke}" stroke-width="{stroke_width}"/>'
        f'</svg>'
    )


def mini_chart(kind: str, data, **kwargs) -> str:
    """Unified entry-point for P2-1 mini charts.

    ``kind`` is one of ``"bar"``/``"mini_bar"``, ``"line"``/``"mini_line"``,
    ``"pie"``/``"mini_pie"``. Remaining kwargs are forwarded to the matching
    generator.
    """
    k = (kind or "").strip().lower()
    if k in ("bar", "mini_bar", "minibar", "sparkbar"):
        return mini_bar(data, **kwargs)
    if k in ("line", "mini_line", "miniline", "sparkline"):
        return mini_line(data, **kwargs)
    if k in ("pie", "mini_pie", "minipie", "sparkpie"):
        return mini_pie(data, **kwargs)
    raise ValueError(f"未知迷你图表类型 {kind!r}；支持 bar/line/pie")
