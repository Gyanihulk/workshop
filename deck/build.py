"""Rebuild the three workshop-opening slides inside the real deck.

    python3 build.py            # reads source/workshop-deck.pptx, writes out/workshop-deck.pptx

Edit the CONTENT section to change wording; the helpers below keep every slide on the
deck's own layouts, fonts (Cambria/Calibri) and palette. Safe to re-run: slides this script
made are found by title and replaced.
"""
from __future__ import annotations

import copy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).parent
SRC = HERE / "source" / "workshop-deck.pptx"
OUT = HERE / "out" / "workshop-deck.pptx"
ICONS = HERE / "icons"

# Palette pulled from the existing deck
NAVY, NAVY_2 = "1E2761", "2A376B"
INK, SLATE, MUTED = "0B1220", "475569", "94A3B8"
PURPLE, CYAN, AMBER = "7C3AED", "22D3EE", "F59E0B"
TINT, WHITE, PANEL, RULE = "EEF2FF", "FFFFFF", "111A33", "334155"
HEAD, BODY = "Cambria", "Calibri"

# Titles of the slides this script owns (used to find and replace them on re-run)
T_INTRO, T_WARM, T_FLOW = "Hi, I'm Adamya", "Warm-up: pick your corner", "From classroom to first client"
LEGACY_TITLES = {T_INTRO, T_WARM, T_FLOW}


# ----------------------------------------------------------------------------- helpers
def rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def shape(slide, kind, x, y, w, h, fill, *, radius=None, shadow=False, name=None):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    s.line.fill.background()
    style = s._element.find(qn("p:style"))
    if style is not None:
        s._element.remove(style)
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    sp = s._element.spPr
    for old in sp.findall(qn("a:effectLst")):
        sp.remove(old)
    eff = etree.SubElement(sp, qn("a:effectLst"))
    if shadow:
        sh = etree.SubElement(
            eff, qn("a:outerShdw"), blurRad="95250", dist="25400", dir="5400000", algn="t", rotWithShape="0"
        )
        clr = etree.SubElement(sh, qn("a:srgbClr"), val=NAVY)
        etree.SubElement(clr, qn("a:alpha"), val="14000")
    if name:
        s.name = name
    return s


def text(slide, x, y, w, h, paras, *, name=None, anchor=MSO_ANCHOR.TOP):
    """paras: list of dicts {t, size, bold, color, font, italic, align, space_after} or a single dict."""
    if isinstance(paras, dict):
        paras = [paras]
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if "space_after" in p:
            para.space_after = Pt(p["space_after"])
        r = para.add_run()
        r.text = p["t"]
        f = r.font
        f.name = p.get("font", BODY)
        f.size = Pt(p.get("size", 14))
        f.bold = p.get("bold", False)
        f.italic = p.get("italic", False)
        f.color.rgb = rgb(p.get("color", INK))
    if name:
        tb.name = name
    return tb


def icon(slide, x, y, d, circle, icon_file, name):
    """Icon centred in a filled circle (diameter d, inches)."""
    shape(slide, MSO_SHAPE.OVAL, x, y, d, d, circle, name=f"{name} circle")
    pad = d * 0.28
    pic = slide.shapes.add_picture(str(ICONS / icon_file), Inches(x + pad), Inches(y + pad), Inches(d - 2 * pad), Inches(d - 2 * pad))
    pic.name = f"{name} icon"


def chip(slide, x, y, w, label, fill=AMBER, color=INK, h=0.36, size=11, name="chip"):
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, radius=0.5, name=name)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name, r.font.size, r.font.bold = BODY, Pt(size), True
    r.font.color.rgb = rgb(color)
    return s


def title_of(slide) -> str:
    return slide.shapes.title.text_frame.text.strip() if slide.shapes.title is not None else ""


def set_title(slide, t, *, size=None, color=None, x=None, y=None, w=None, h=None):
    ph = slide.shapes.title
    ph.text_frame.text = t
    r = ph.text_frame.paragraphs[0].runs[0]
    r.font.name, r.font.bold = HEAD, True
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = rgb(color)
    if x is not None:
        ph.left, ph.top, ph.width, ph.height = Inches(x), Inches(y), Inches(w), Inches(h)
    return ph


def delete_slide(prs, slide):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        if prs.part.related_part(sld.rId) is slide.part:
            prs.part.drop_rel(sld.rId)
            ids.remove(sld)
            return


def move_slide(prs, slide, index):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        if prs.part.related_part(sld.rId) is slide.part:
            ids.remove(sld)
            ids.insert(index, sld)
            return


def add_slide_number(slide, donor):
    """Copy the slide-number placeholder from a slide that already has one."""
    for sp in donor.shapes:
        if sp.is_placeholder and sp.placeholder_format.type == 13:  # SLIDE_NUMBER
            slide.shapes._spTree.append(copy.deepcopy(sp._element))
            return


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body.strip()


# ----------------------------------------------------------------------------- slides
def build_intro(prs, layout):
    s = prs.slides.add_slide(layout)
    # decorative circle, same motif as the cover slide
    shape(s, MSO_SHAPE.OVAL, 10.9, -2.3, 5.2, 5.2, PURPLE, name="Decor circle")
    text(s, 0.9, 0.85, 6.0, 0.3, dict(t="ABOUT YOUR SPEAKER", size=14, bold=True, color=CYAN), name="Eyebrow")
    set_title(s, T_INTRO, size=54, color=WHITE, x=0.8, y=1.25, w=6.3, h=1.1)
    text(
        s, 0.9, 2.5, 6.0, 1.0,
        dict(t="A working engineer, not a professional speaker. This is what I wish I had known in my final year.", size=18, color="CBD5E1"),
        name="Subtitle",
    )
    rows = [
        ("rocket_light_FFFFFF.png", "10 products shipped solo", "From idea to real users, outside the day job"),
        ("project_21D2ED.png", "Backend Engineer, Hero MotoCorp", "Kafka, AWS and GCP, real production traffic"),
        ("client_F49D0A.png", "Co-founder, Adhian Gurus", "A zero-commission tutoring marketplace"),
    ]
    for i, (ic, head, sub) in enumerate(rows):
        y = 3.75 + i * 0.88
        icon(s, 0.9, y, 0.66, "162238", ic, f"Stat {i + 1}")
        text(s, 1.8, y + 0.02, 5.2, 0.35, dict(t=head, size=17, bold=True, color=WHITE, font=HEAD), name=f"Stat {i + 1} head")
        text(s, 1.8, y + 0.36, 5.2, 0.3, dict(t=sub, size=13, color=MUTED), name=f"Stat {i + 1} sub")
    text(
        s, 0.9, 6.55, 6.4, 0.4,
        dict(t="A few years ago I was in your seat. Today is the path I actually took.", size=15, italic=True, color="E2E8F0", font=HEAD),
        name="Closing line",
    )

    # timeline panel
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 7.7, 0.9, 5.0, 5.75, PANEL, radius=0.05, name="Timeline panel")
    text(s, 8.15, 1.25, 4.2, 0.3, dict(t="THE PATH", size=12, bold=True, color=CYAN), name="Timeline label")
    shape(s, MSO_SHAPE.RECTANGLE, 8.31, 2.15, 0.04, 3.15, RULE, name="Timeline rule")
    steps = [
        ("2021", "ITDA", "Training Desk Analyst"),
        ("2021–22", "Capline Services", "Associate"),
        ("2022–24", "Appfoster", "Software Engineer II"),
        ("2024–Now", "Hero MotoCorp", "Backend Engineer"),
    ]
    for i, (yr, org, role) in enumerate(steps):
        y = 1.95 + i * 1.05
        dot = AMBER if i == len(steps) - 1 else PURPLE
        shape(s, MSO_SHAPE.OVAL, 8.15, y + 0.12, 0.36, 0.36, dot, name=f"Dot {i + 1}")
        text(s, 8.8, y, 3.6, 0.25, dict(t=yr, size=12, bold=True, color=CYAN), name=f"Year {i + 1}")
        text(s, 8.8, y + 0.25, 3.6, 0.35, dict(t=org, size=18, bold=True, color=WHITE, font=HEAD), name=f"Org {i + 1}")
        text(s, 8.8, y + 0.62, 3.6, 0.3, dict(t=role, size=13, color=MUTED), name=f"Role {i + 1}")
    text(
        s, 8.15, 6.12, 4.3, 0.4,
        dict(t="B.Tech CSE (2019) · MBA Finance & HR, Gurukul Kangri, Haridwar", size=11, color=MUTED),
        name="Education",
    )
    notes(
        s,
        """About 60 seconds. Point first: "I'll show you a path, not a pitch."
Name the Adhian Gurus co-founder role plainly, once, so nobody is surprised when it shows up as a case study later. Do not sell it.
Facts on this slide come from your own earlier slide ("10 products", the four jobs); confirm they are still right before presenting.""",
    )
    return s


def build_warmup(prs, layout, donor):
    s = prs.slides.add_slide(layout)
    set_title(s, T_WARM)
    text(
        s, 0.7, 1.3, 12.0, 0.4,
        dict(t="Which one do you most want help with today? Walk to that corner. There is no wrong answer.", size=16, color=SLATE),
        name="Instruction",
    )
    cells = [
        ("resume", "resume_7B39EC.png", "Resume", "Bullets, layout and the 7-second scan", "FRONT LEFT"),
        ("dsa", "dsa_7B39EC.png", "DSA patterns", "What to practice, and how to revisit it", "FRONT RIGHT"),
        ("project", "project_21D2ED.png", "Building a project", "Pick it, scope it, ship it", "BACK LEFT"),
        ("client", "client_F49D0A.png", "Getting a first client", "Network, pitch and pricing", "BACK RIGHT"),
    ]
    W, H, GAP, X0, Y0 = 5.95, 1.8, 0.25, 0.6, 1.95
    for i, (key, ic, head, sub, corner) in enumerate(cells):
        col, row = i % 2, i // 2
        x, y = X0 + col * (W + GAP), Y0 + row * (H + GAP)
        dark = (col + row) % 2 == 1  # checkerboard
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, W, H, NAVY if dark else TINT, radius=0.08, shadow=True, name=f"Corner card: {head}")
        icon(s, x + 0.4, y + 0.45, 0.9, NAVY_2 if dark else WHITE, ic, head)
        text(s, x + 1.6, y + 0.28, W - 1.9, 0.5, dict(t=head, size=24, bold=True, font=HEAD, color=WHITE if dark else INK), name=f"{head} title")
        text(s, x + 1.6, y + 0.85, W - 1.9, 0.3, dict(t=sub, size=14, color="CBD5E1" if dark else SLATE), name=f"{head} sub")
        chip(s, x + 1.6, y + 1.28, 1.5, corner, name=f"{head} corner")
    # prompt band
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 6.05, 12.15, 0.8, NAVY, radius=0.18, name="Prompt band")
    chip(s, 0.9, 6.27, 1.1, "2 MIN", h=0.36, size=13, name="Timer chip")
    text(
        s, 2.25, 6.12, 10.2, 0.66,
        dict(t="In your corner: what have you tried, and what went wrong? Then one person per corner shares one sentence.", size=16, color=WHITE),
        name="Prompt",
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_slide_number(s, donor)
    notes(
        s,
        """6 minutes total.
1. Point to the four corners and say: "Stand in the corner for the thing you're most stuck on. There is no wrong answer." (30 s)
2. Everyone walks. Let them settle. (1 min)
3. In each corner, groups of 3 to 4 answer the question on the slide. (2 min)
4. One person per corner shares one sentence. Write each corner's headcount on the board. (2 min)
5. Close: "Today runs in this order. Watch for your corner." Then go to the flow slide. (30 s)
Room packed: raise hands for each topic instead of walking. Online: run it as a poll in chat with the same four options.
Optional second round (3 min): pairs swap a resume, the partner looks for 7 seconds and says what stood out. Reuse it at "What a recruiter actually scans for".""",
    )
    return s


def build_flow(prs, layout, donor):
    s = prs.slides.add_slide(layout)
    set_title(s, T_FLOW)
    text(s, 0.7, 1.3, 12.0, 0.4, dict(t="Four steps. Each one builds on the one before.", size=16, color=SLATE), name="Subtitle")
    steps = [
        ("PART 1", "Resume", "One page. Numbers, not adjectives. A live link.", "resume_7B39EC.png"),
        ("PART 2", "DSA patterns", "About 15 patterns, revisited, beat raw problem count.", "dsa_7B39EC.png"),
        ("PART 3", "Ship a project", "Pick a real annoyance, scope it to a weekend, deploy it.", "project_21D2ED.png"),
        ("PART 4", "First client", "Start with your network, pitch their problem, ask for a testimonial.", "client_F49D0A.png"),
    ]
    COL, D, Y = 3.1, 1.2, 2.25
    first_c, last_c = 0.6 + D / 2, 0.6 + 3 * COL + D / 2
    shape(s, MSO_SHAPE.RECTANGLE, first_c, Y + D / 2 - 0.03, last_c - first_c, 0.06, "CBD5E1", name="Path line")
    for i, (tag, head, body, ic) in enumerate(steps):
        x = 0.6 + i * COL
        if i < 3:  # arrow between nodes
            tri = shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, x + COL / 2 + D / 4, Y + D / 2 - 0.14, 0.28, 0.28, "94A3B8", name=f"Arrow {i + 1}")
            tri.rotation = 90
        icon(s, x, Y, D, NAVY, ic, head)
        chip(s, x + D - 0.3, Y - 0.1, 0.42, str(i + 1), h=0.42, size=14, name=f"Step number {i + 1}")
        text(s, x, Y + D + 0.35, 2.7, 0.25, dict(t=tag, size=12, bold=True, color=PURPLE), name=f"{head} tag")
        text(s, x, Y + D + 0.65, 2.7, 0.45, dict(t=head, size=22, bold=True, font=HEAD), name=f"{head} title")
        text(s, x, Y + D + 1.2, 2.65, 1.2, dict(t=body, size=16, color=SLATE), name=f"{head} body")
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 6.0, 12.15, 0.8, TINT, radius=0.18, shadow=True, name="Optional band")
    chip(s, 0.9, 6.22, 1.35, "OPTIONAL", h=0.36, size=12, name="Optional chip")
    text(
        s, 2.5, 6.07, 10.0, 0.66,
        dict(t="Mock interviews for anyone who wants practice. Opt in, no pressure.", size=16, color=INK),
        name="Optional text",
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_slide_number(s, donor)
    notes(
        s,
        """Use this as the map for the session and as the discussion slide at the end.
Every number here already appears later in the deck (about 15 patterns, weekend scope).
Do not mention days or the tutoring marketplace on this slide.""",
    )
    return s


# ----------------------------------------------------------------------------- main
def main():
    prs = Presentation(SRC)
    # 1. drop earlier versions of these slides
    for sl in list(prs.slides):
        texts = {sh.text_frame.text.strip() for sh in sl.shapes if sh.has_text_frame}
        if title_of(sl) in LEGACY_TITLES or texts & LEGACY_TITLES:
            delete_slide(prs, sl)

    # python-pptx names the next slide part after the slide count, so renumber after deleting
    prs.part.rename_slide_parts([sld.rId for sld in prs.slides._sldIdLst])

    by_name = {l.name: l for l in prs.slide_layouts}
    dark, light = by_name["TitleDark"], by_name["ContentLight"]
    donor = next(sl for sl in prs.slides if title_of(sl) == "What we'll cover today")

    intro = build_intro(prs, dark)
    warm = build_warmup(prs, light, donor)
    flow = build_flow(prs, light, donor)

    # 2. position: intro + warm-up right after the cover, flow right after "What we'll cover today"
    move_slide(prs, intro, 1)
    move_slide(prs, warm, 2)
    titles = [title_of(sl) for sl in prs.slides]
    move_slide(prs, flow, titles.index("What we'll cover today") + 1)

    OUT.parent.mkdir(exist_ok=True)
    prs.save(OUT)
    print("wrote", OUT, "-", len(prs.slides), "slides")


if __name__ == "__main__":
    main()
