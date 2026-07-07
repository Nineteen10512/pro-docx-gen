"""LaTeX → OMML (Office Math ML) converter for DOCX v1.3 (DOCX-P0-1).

Uses ``latex2mathml`` for LaTeX → MathML, then converts MathML → OMML via a
hand-written XSLT-like walker (no external XSLT needed; MML2OMML.XSL is not
available on all platforms). Supports the most common constructs used in
academic and teaching documents:

- identifiers, numbers, operators
- sub/superscript (msub/msup/msubsup)
- fractions (mfrac)
- roots/sqrt (msqrt/mroot)
- Greek letters (auto-convert to Unicode math symbols)
- operators: ∑, ∫, √, ≠, ≤, ≥, ±, ×, ÷, ∞, etc. (via latex2mathml operator table)
- matrices (mtable)
- cases (mfenced + mtable)
- fences (parentheses, brackets, braces)
- accents/over (hat, bar, dot, vec, etc.)
- integrals, summation, product with limits (munder/mover/munderover)

Inline math uses ``m:oMath``; display math uses ``m:oMathPara`` with
(optional) right-aligned equation number in caption.

@since v1.3.0 (DOCX-P0-1)
"""
from __future__ import annotations

from typing import Optional

from lxml import etree

# Namespaces
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

M = "{%s}" % M_NS
W = "{%s}" % W_NS


# ---------------------------------------------------------------------------
# LaTeX → MathML (via latex2mathml)
# ---------------------------------------------------------------------------

def _latex_to_mathml(latex: str, display: str = "inline") -> str:
    import latex2mathml.converter as conv
    # latex2mathml.converter.convert returns a MathML string (xml)
    kwargs = {}
    if display == "block":
        kwargs["display"] = "block"
    mml_str = conv.convert(latex, **kwargs)
    # Normalize: ensure wrapper tag is <math ...>
    if not mml_str.strip().startswith("<math"):
        mml_str = f'<math xmlns="http://www.w3.org/1998/Math/MathML">{mml_str}</math>'
    return mml_str


# ---------------------------------------------------------------------------
# MathML → OMML walker
# ---------------------------------------------------------------------------

# Map of MathML tag → OMML builder function.
# Each builder is called with (xml_element, builder) and returns an OMML element.

# Operator substitution table (mathml entity / latex command → OMML-friendly char)
_OPERATOR_MAP = {
    "\u2061": "",      # &ApplyFunction;
    "\u2062": "",      # invisible times
    "\u2063": "",      # separator
    "\u2064": "",      # invisible plus
    "\u00b7": "\u00b7",
    "\u00d7": "\u00d7",
    "\u00f7": "\u00f7",
    "\u2260": "\u2260",
    "\u2264": "\u2264",
    "\u2265": "\u2265",
    "\u00b1": "\u00b1",
    "\u221e": "\u221e",
    "\u2211": "\u2211",
    "\u222b": "\u222b",
    "\u221a": "\u221a",
    "\u2202": "\u2202",
    "\u0394": "\u0394",
    "\u03b1": "\u03b1",
    "\u03b2": "\u03b2",
    "\u03b3": "\u03b3",
    "\u03b8": "\u03b8",
    "\u03c0": "\u03c0",
    "\u03c3": "\u03c3",
    "\u03c6": "\u03c6",
    "\u03a9": "\u03a9",
}


class _Mml2Omml:
    """Converter state for a single equation."""

    def __init__(self):
        self._seen_rpr = False

    def convert(self, mml_str: str) -> etree._Element:
        try:
            root = etree.fromstring(mml_str.encode("utf-8"))
        except etree.XMLSyntaxError as e:
            # Fallback: wrap in m:r with raw LaTeX text
            oMath = etree.Element(M + "oMath", nsmap={"m": M_NS})
            r = etree.SubElement(oMath, M + "r")
            t = etree.SubElement(r, M + "t")
            t.text = mml_str
            return oMath
        return self._build_oMath_from_math(root)

    def _build_oMath_from_math(self, math_el: etree._Element) -> etree._Element:
        oMath = etree.Element(M + "oMath", nsmap={"m": M_NS})
        for child in math_el:
            self._build_into(oMath, child)
        return oMath

    # ---- dispatch --------------------------------------------------------

    def _build_into(self, parent: etree._Element, el: etree._Element):
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        handler = getattr(self, f"_m_{tag}", self._m_default)
        handler(parent, el)

    def _m_default(self, parent, el):
        # Unknown tag: recurse into children; if text, emit run
        if el.text and el.text.strip():
            self._emit_text(parent, el.text)
        for ch in el:
            self._build_into(parent, ch)
            if ch.tail and ch.tail.strip():
                self._emit_text(parent, ch.tail)

    def _emit_text(self, parent, text: str, style: str = "p"):
        """Emit m:r with m:t containing the given text."""
        text = _OPERATOR_MAP.get(text, text)
        if not text:
            return
        r = etree.SubElement(parent, M + "r")
        if style != "p":
            rPr = etree.SubElement(r, M + "rPr")
            if style == "b":
                sty = etree.SubElement(rPr, M + "sty"); sty.set("val", "b")
            elif style == "i":
                sty = etree.SubElement(rPr, M + "sty"); sty.set("val", "i")
            elif style == "bi":
                sty = etree.SubElement(rPr, M + "sty"); sty.set("val", "bi")
            # upright style for operators/numerals
            if style == "p":
                sty = etree.SubElement(rPr, M + "sty"); sty.set("val", "p")
        t = etree.SubElement(r, M + "t")
        # Preserve spaces
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text

    # ---- common MathML elements ------------------------------------------

    def _m_mi(self, p, el):
        # identifier — typically italic for single letters, upright for multi-char
        text = (el.text or "").strip()
        if not text:
            for ch in el:
                self._build_into(p, ch)
            return
        # Check for mathvariant attribute
        variant = el.get("mathvariant", "")
        style = "i"  # default italic for single letter
        if len(text) > 1 or variant == "normal" or variant == "upright":
            style = "p"
        elif variant == "bold":
            style = "b"
        elif variant == "bold-italic":
            style = "bi"
        # Greek letters — use Unicode directly (don't change style)
        self._emit_text(p, text, style=style)

    def _m_mn(self, p, el):
        self._emit_text(p, (el.text or "").strip(), style="p")

    def _m_mo(self, p, el):
        text = (el.text or "").strip()
        text = _OPERATOR_MAP.get(text, text)
        if not text:
            return
        # operator; usually upright
        r = etree.SubElement(p, M + "r")
        rPr = etree.SubElement(r, M + "rPr")
        sty = etree.SubElement(rPr, M + "sty"); sty.set("val", "p")
        t = etree.SubElement(r, M + "t"); t.text = text

    def _m_mtext(self, p, el):
        self._emit_text(p, (el.text or "").strip(), style="p")

    def _m_mspace(self, p, el):
        # space as small whitespace
        self._emit_text(p, " ", style="p")

    def _m_msub(self, p, el):
        # m:sub sSub base sub
        children = [ch for ch in el if not ch.tag.endswith("}") is False or True]
        # we iterate non-text children
        kids = [c for c in el]
        if len(kids) >= 2:
            sSub = etree.SubElement(p, M + "sSub")
            e = etree.SubElement(sSub, M + "e"); self._build_into(e, kids[0])
            sub = etree.SubElement(sSub, M + "sub"); self._build_into(sub, kids[1])

    def _m_msup(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 2:
            sSup = etree.SubElement(p, M + "sSup")
            e = etree.SubElement(sSup, M + "e"); self._build_into(e, kids[0])
            sup = etree.SubElement(sSup, M + "sup"); self._build_into(sup, kids[1])

    def _m_msubsup(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 3:
            sSubSup = etree.SubElement(p, M + "sSubSup")
            e = etree.SubElement(sSubSup, M + "e"); self._build_into(e, kids[0])
            sub = etree.SubElement(sSubSup, M + "sub"); self._build_into(sub, kids[1])
            sup = etree.SubElement(sSubSup, M + "sup"); self._build_into(sup, kids[2])

    def _m_mfrac(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 2:
            f = etree.SubElement(p, M + "f")
            fPr = etree.SubElement(f, M + "fPr")
            # bar type default (use frac bar)
            num = etree.SubElement(f, M + "num"); self._build_into(num, kids[0])
            den = etree.SubElement(f, M + "den"); self._build_into(den, kids[1])

    def _m_msqrt(self, p, el):
        rad = etree.SubElement(p, M + "rad")
        radPr = etree.SubElement(rad, M + "radPr")
        degHide = etree.SubElement(radPr, M + "degHide"); degHide.set("val", "1")
        deg = etree.SubElement(rad, M + "deg")
        e = etree.SubElement(rad, M + "e")
        for ch in el: self._build_into(e, ch)

    def _m_mroot(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 2:
            rad = etree.SubElement(p, M + "rad")
            deg = etree.SubElement(rad, M + "deg"); self._build_into(deg, kids[1])
            e = etree.SubElement(rad, M + "e"); self._build_into(e, kids[0])
        else:
            self._m_msqrt(p, el)

    def _m_mfenced(self, p, el):
        open_c = el.get("open", "(")
        close_c = el.get("close", ")")
        separators = el.get("separators", ",")
        # OMML uses m:d (delimiter element)
        d = etree.SubElement(p, M + "d")
        dPr = etree.SubElement(d, M + "dPr")
        begChr = etree.SubElement(dPr, M + "begChr"); begChr.set("val", open_c or "")
        endChr = etree.SubElement(dPr, M + "endChr"); endChr.set("val", close_c or "")
        if separators:
            sepChr = etree.SubElement(dPr, M + "sepChr"); sepChr.set("val", separators[0])
        kids = [c for c in el]
        for k in kids:
            if k.tag.endswith("}mrow") or k.tag.endswith("}mrow".lower()):
                e = etree.SubElement(d, M + "e"); self._build_into(e, k)
            else:
                e = etree.SubElement(d, M + "e"); self._build_into(e, k)

    def _m_mrow(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_mstyle(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_merror(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_mpadded(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_mphantom(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_menclose(self, p, el):
        notation = el.get("notation", "longdiv")
        # box/strike/etc. — fall back to grouping under m:d for parenthesis notation
        if notation == "phasorangle":
            # Not common — just emit content
            for ch in el: self._build_into(p, ch)
        else:
            # treat as parentheses
            d = etree.SubElement(p, M + "d")
            e = etree.SubElement(d, M + "e")
            for ch in el: self._build_into(e, ch)

    def _m_munder(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 2:
            limLow = etree.SubElement(p, M + "nary")
            # check if first child is large operator (∑ ∫ ∏)
            op_text = "".join(kids[0].itertext()).strip()
            naryPr = etree.SubElement(limLow, M + "naryPr")
            chr_el = etree.SubElement(naryPr, M + "chr"); chr_el.set("val", op_text or "∑")
            limLoc = etree.SubElement(naryPr, M + "limLoc"); limLoc.set("val", "und")
            sub = etree.SubElement(limLow, M + "sub"); self._build_into(sub, kids[1])
            sup = etree.SubElement(limLow, M + "sup")
            e = etree.SubElement(limLow, M + "e")
            # no base
            etree.SubElement(e, M + "r")

    def _m_mover(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 2:
            # If upper child is a ^-style bar/hat, use m:acc (accent)
            upper_text = "".join(kids[1].itertext()).strip()
            if upper_text in ("^", "ˆ", "‾", "¯", "→", ".", "..", "~", "⃗"):
                acc = etree.SubElement(p, M + "acc")
                accPr = etree.SubElement(acc, M + "accPr")
                chr_map = {"^":"̂","‾":"̅","¯":"̅","→":"⃗",".":"̇","..":"̈","~":"̃","ˆ":"̂","⃗":"⃗"}
                ch_el = etree.SubElement(accPr, M + "chr"); ch_el.set("val", chr_map.get(upper_text, upper_text))
                e = etree.SubElement(acc, M + "e"); self._build_into(e, kids[0])
            else:
                bar = etree.SubElement(p, M + "bar"); self._build_into(bar, kids[0])

    def _m_munderover(self, p, el):
        kids = [c for c in el]
        if len(kids) >= 3:
            nary = etree.SubElement(p, M + "nary")
            naryPr = etree.SubElement(nary, M + "naryPr")
            op_text = "".join(kids[0].itertext()).strip()
            chr_el = etree.SubElement(naryPr, M + "chr"); chr_el.set("val", op_text or "∑")
            # Determine limLoc: for sum/prod/∏ use undOvr; for integral use subSup
            if op_text in ("∫", "∬", "∭"):
                limLoc = etree.SubElement(naryPr, M + "limLoc"); limLoc.set("val", "subSup")
            else:
                limLoc = etree.SubElement(naryPr, M + "limLoc"); limLoc.set("val", "undOvr")
            sub = etree.SubElement(nary, M + "sub"); self._build_into(sub, kids[1])
            sup = etree.SubElement(nary, M + "sup"); self._build_into(sup, kids[2])
            e = etree.SubElement(nary, M + "e")
            # empty base (the operator renders as chr)
            r = etree.SubElement(e, M + "r"); etree.SubElement(r, M + "t")

    def _m_mtable(self, p, el):
        # Matrix: OMML m:m
        mMat = etree.SubElement(p, M + "m")
        mPr = etree.SubElement(mMat, M + "mPr")
        # Defaults: parentheses around matrix
        # Detect cases (first child is mfenced with open='{')
        baseEl = el.getparent() if hasattr(el, "getparent") else None
        for tr in el:
            tag = tr.tag.split("}")[-1] if "}" in tr.tag else tr.tag
            if tag != "mtr" and tag != "mlabeledtr":
                continue
            mr = etree.SubElement(mMat, M + "mr")
            cells = tr.findall(".//{http://www.w3.org/1998/Math/MathML}mtd")
            if not cells:
                cells = [c for c in tr if c.tag.endswith("}mtd")]
            if not cells:
                cells = [tr]
            for td in cells:
                e = etree.SubElement(mr, M + "e")
                self._build_into(e, td)

    # mtr/mtd handled via recursion (only exist inside mtable)
    def _m_mtr(self, p, el):
        for ch in el: self._build_into(p, ch)
    def _m_mtd(self, p, el):
        for ch in el: self._build_into(p, ch)
    def _m_mlabeledtr(self, p, el):
        for ch in el: self._build_into(p, ch)

    def _m_mmultiscripts(self, p, el):
        # Complex pre/post scripts — fall back to sub/sup reading postscripts
        kids = [c for c in el if not (c.tag.endswith("}mprescripts"))]
        # Naive: treat first as base, subsequent pairs as sub/sup
        if kids:
            base = kids[0]
            rest = kids[1:]
            for r_ in rest:
                np = etree.SubElement(p, M + "sSup")
                e = etree.SubElement(np, M + "e")
                if r_ == rest[0]:
                    self._build_into(e, base)
                sup = etree.SubElement(np, M + "sup"); self._build_into(sup, r_)
                base_for_next = np
                # This is rough — sufficient for basic tensors


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def latex_to_omml(latex: str, display: str = "inline") -> etree._Element:
    """Convert LaTeX math string to an ``m:oMath`` (inline) or
    ``m:oMathPara`` (block) element ready to insert into a paragraph.

    Returns an lxml element.
    """
    mml = _latex_to_mathml(latex, display=display)
    conv = _Mml2Omml()
    oMath = conv.convert(mml)
    if display == "block":
        para = etree.Element(M + "oMathPara", nsmap={"m": M_NS})
        oMathParaPr = etree.SubElement(para, M + "oMathParaPr")
        jc = etree.SubElement(oMathParaPr, M + "jc"); jc.set("val", "center")
        para.append(oMath)
        return para
    return oMath


def append_equation_to_paragraph(paragraph, latex: str, display: str = "inline",
                                 caption: Optional[str] = None,
                                 cn_font: str = "宋体",
                                 caption_size_pt: float = 10.5):
    """Append an OMML equation to a python-docx paragraph.

    For display=block: equation is centered; if caption is provided, a
    right-aligned tab is used to place (3-1) style numbering on the right.
    """
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    if display == "block":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Hidden source-latex run for round-trippability (searchable in XML, invisible in Word)
        _add_hidden_latex_run(paragraph, latex)
        omml = latex_to_omml(latex, display="block")
        # oMathPara wrapper: insert directly into paragraph after pPr
        # Note: a w:p can contain m:oMathPara as direct child.
        pPr = paragraph._p.find(_qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(omml)
        else:
            paragraph._p.insert(0, omml)
        if caption:
            # Append tab + caption in right-aligned form via m:r? We use the
            # simpler approach: append a tab-stop right-aligned and caption run
            # on the same paragraph.
            tab = OxmlElement("w:r"); t = OxmlElement("w:tab"); tab.append(t)
            # add right tab stop
            if pPr is None:
                pPr = OxmlElement("w:pPr"); paragraph._p.insert(0, pPr)
            tabs = pPr.find(_qn("w:tabs"))
            if tabs is None:
                tabs = OxmlElement("w:tabs"); pPr.append(tabs)
            tab_el = OxmlElement("w:tab")
            tab_el.set(_qn("w:val"), "right")
            tab_el.set(_qn("w:pos"), "9000")
            tabs.append(tab_el)
            paragraph._p.append(tab)
            run = paragraph.add_run(caption)
            run.font.size = Pt(caption_size_pt)
            run.font.name = "Times New Roman"
            # Set East-Asian font
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(_qn("w:rFonts"))
            if rFonts is None:
                from lxml import etree as _et
                rFonts = _et.SubElement(rPr, _qn("w:rFonts"))
                rPr.insert(0, rFonts)
            rFonts.set(_qn("w:ascii"), "Times New Roman")
            rFonts.set(_qn("w:hAnsi"), "Times New Roman")
            rFonts.set(_qn("w:eastAsia"), cn_font)
            rFonts.set(_qn("w:cs"), "Times New Roman")
    else:
        # Inline: prepend a hidden run containing raw latex for traceability
        _add_hidden_latex_run(paragraph, latex)
        omml = latex_to_omml(latex, display="inline")
        paragraph._p.append(omml)


def _add_hidden_latex_run(paragraph, latex: str):
    """Append a vanish (hidden) w:r containing the raw LaTeX source,
    so that the equation source is preserved in XML for search/audit
    but not rendered in Word."""
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    vanish = OxmlElement("w:vanish")
    rPr.append(vanish)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = latex
    t.set(_qn("xml:space"), "preserve")
    r.append(t)
    paragraph._p.append(r)


__all__ = ["latex_to_omml", "append_equation_to_paragraph"]
