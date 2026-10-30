# -*- coding: utf-8 -*-
"""DMart — Trolley-Rest Pilot Pitch (10 slides) — earthy palette (Color pallette.pdf)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# PALETTE (exact hex from Color pallette.pdf + derived tints)
CREAM     = RGBColor(0xF2, 0xEF, 0xE9)
DARK      = RGBColor(0x3A, 0x4C, 0x40)
SAGE      = RGBColor(0x8B, 0xA8, 0x92)
CLAY      = RGBColor(0xC4, 0x69, 0x4A)
SAGE_PALE = RGBColor(0xE6, 0xEC, 0xE8)
CLAY_PALE = RGBColor(0xF4, 0xE7, 0xE1)
INK       = RGBColor(0x2B, 0x38, 0x30)
GREY      = RGBColor(0x6E, 0x7A, 0x72)
LINE      = RGBColor(0xD8, 0xDE, 0xD9)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

HEAD = "Trebuchet MS"
BODY = "Calibri"


def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.0,
         shape=MSO_SHAPE.RECTANGLE, rounding=None):
    sp = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def oval(slide, l, t, w, h, fill=None, line=None, line_w=1.0):
    return rect(slide, l, t, w, h, fill, line, line_w, MSO_SHAPE.OVAL)


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    return tb, tf


def block(slide, l, t, w, h, paras, anchor=MSO_ANCHOR.TOP):
    tb, tf = textbox(slide, l, t, w, h, anchor=anchor)
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        para.space_after = Pt(p.get("space_after", 6))
        para.space_before = Pt(p.get("space_before", 0))
        para.line_spacing = p.get("line", 1.0)
        run = para.add_run()
        run.text = (("\u25aa  " if p.get("bullet") else "") + p["text"])
        run.font.size = Pt(p["size"])
        run.font.bold = p.get("bold", False)
        run.font.color.rgb = p["color"]
        run.font.name = p.get("font", BODY)
    return tb


def needed(slide, l, t, text="[NEEDED]"):
    w = 0.16 + 0.092 * len(text)
    rect(slide, l, t, w, 0.34, fill=CLAY_PALE, line=CLAY, line_w=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(slide, l, t + 0.02, w, 0.30,
          [{"text": text, "size": 11, "color": CLAY, "bold": True, "font": HEAD,
            "align": PP_ALIGN.CENTER}])
    return w


def header(slide, eyebrow, headline, hcolor=INK, hsize=29):
    rect(slide, 0, 0, 13.333, 0.16, DARK)
    rect(slide, 0.7, 0.46, 0.30, 0.30, CLAY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(slide, 1.12, 0.47, 10.5, 0.32,
          [{"text": eyebrow, "size": 12, "color": DARK, "bold": True, "font": HEAD}])
    block(slide, 0.7, 0.92, 12.0, 1.55,
          [{"text": headline, "size": hsize, "color": hcolor, "bold": True,
            "font": HEAD, "line": 1.02}])


def footer(slide, n):
    rect(slide, 0.7, 6.96, 11.93, 0.015, LINE)
    block(slide, 0.7, 7.04, 7.5, 0.35,
          [{"text": "DMart  \u00b7  Trolley-Rest Pilot Pitch", "size": 9, "color": GREY, "font": BODY}])
    block(slide, 4.6, 7.04, 5.3, 0.35,
          [{"text": "Source: Group 7 HCD study \u00b7 DMart + Star Bazaar observations",
            "size": 9, "color": GREY, "font": BODY, "align": PP_ALIGN.CENTER}])
    block(slide, 11.2, 7.04, 1.4, 0.35,
          [{"text": f"{n:02d} / 10", "size": 9, "color": DARK, "bold": True,
            "font": HEAD, "align": PP_ALIGN.RIGHT}])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = CREAM
    return s


def journey(slide, top, items, hi):
    n = len(items)
    total = 11.9
    gap = 0.34
    nw = (total - gap * (n - 1)) / n
    x = 0.7
    for i, (label, sub) in enumerate(items):
        active = i in hi
        fill = DARK if active else WHITE
        line = DARK if active else LINE
        rect(slide, x, top, nw, 0.86, fill=fill, line=line, line_w=1.25,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        block(slide, x, top + 0.12, nw, 0.62,
              [{"text": label, "size": 13, "bold": True,
                "color": WHITE if active else INK, "font": HEAD, "align": PP_ALIGN.CENTER},
               {"text": sub, "size": 9.5,
                "color": WHITE if active else GREY, "font": BODY, "align": PP_ALIGN.CENTER}])
        if i < n - 1:
            rect(slide, x + nw + 0.02, top + 0.30, gap - 0.04, 0.26,
                 fill=SAGE, shape=MSO_SHAPE.CHEVRON)
        x += nw + gap


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# SLIDE 1
s = new_slide(prs)
header(s, "THE HOOK  \u00b7  01",
       "Your highest-value shoppers are the ones standing in line with nowhere to sit.")
block(s, 0.7, 2.7, 7.4, 3.6, [
    {"text": "THE MAIN POINT", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 4},
    {"text": "Give tired shoppers a rest on their own trolley, and the billing-queue wait becomes a reason to come back \u2014 not a reason to rush out.",
     "size": 17, "color": INK, "font": BODY, "line": 1.12, "space_after": 18},
    {"text": "THE SHOPPER TRUTH", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 4},
    {"text": "When asked the biggest in-store problem, shoppers did not complain about price or stock \u2014 they complained about having nowhere to sit or rest.",
     "size": 15, "color": GREY, "font": BODY, "line": 1.12},
])
oval(s, 9.45, 2.95, 3.05, 3.05, fill=CLAY)
block(s, 9.45, 3.55, 3.05, 1.5,
      [{"text": "75%", "size": 86, "bold": True, "color": WHITE, "font": HEAD, "align": PP_ALIGN.CENTER}])
block(s, 9.45, 5.05, 3.05, 1.0,
      [{"text": "of surveyed shoppers say\nNO REST is a grave problem",
        "size": 13, "bold": True, "color": WHITE, "font": BODY, "align": PP_ALIGN.CENTER, "line": 1.05}])
block(s, 9.45, 6.05, 3.05, 0.4,
      [{"text": "SURVEY SIGNAL", "size": 10, "color": CLAY, "bold": True, "font": HEAD, "align": PP_ALIGN.CENTER}])
footer(s, 1)
notes(s, "Open with the tension: DMart wins on price, but the moment that sticks with a tired "
          "shopper is the billing queue. 75% of the study's surveyed shoppers named 'nowhere to "
          "rest' as a grave problem \u2014 bigger than price or stock complaints. That is a "
          "loyalty lever DMart is not yet pulling. Set up the rest of the deck as the fix.")

# SLIDE 2
s = new_slide(prs)
header(s, "THE PROBLEM  \u00b7  02",
       "The shoppers who spend the most time and money are exactly the ones left standing.")
block(s, 0.7, 2.7, 6.2, 3.8, [
    {"text": "Older, pregnant and diverse shoppers face long billing queues in the tight space around the counters \u2014 with no resting place. The study ties this to exhaustion and a poorer store experience.",
     "size": 15.5, "color": INK, "font": BODY, "line": 1.18, "space_after": 14, "bullet": True},
    {"text": "The empathy map is blunt: at some point these shoppers 'just want to bill and exit.' Tired = shorter browse = fewer unplanned picks.",
     "size": 15.5, "color": INK, "font": BODY, "line": 1.18, "space_after": 14, "bullet": True},
    {"text": "It is a marketing problem, not just a comfort one: a worn-out exit hurts repeat intent \u2014 the metric DMart lives on.",
     "size": 15.5, "color": INK, "font": BODY, "line": 1.18, "bullet": True},
])
def databar(top, label, pct, val, bar_fill):
    block(s, 7.2, top, 5.4, 0.3, [{"text": label, "size": 12.5, "color": INK, "bold": True, "font": HEAD}])
    rect(s, 7.2, top + 0.36, 5.4, 0.5, fill=SAGE_PALE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 7.2, top + 0.36, 5.4 * pct / 100.0, 0.5, fill=bar_fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, 7.2 + 5.4 * pct / 100.0 - 1.0, top + 0.40, 1.0, 0.42,
          [{"text": f"{val}%", "size": 14, "bold": True, "color": WHITE, "font": HEAD, "align": PP_ALIGN.RIGHT}])
databar(2.85, "No resting / sitting  \u2014  grave problem", 75, 75, DARK)
databar(4.35, "Loose-grain section problems", 64, 64, SAGE)
block(s, 7.2, 5.55, 5.4, 1.0,
      [{"text": "n = ~50 survey responses (college, family & relatives). Representativeness of DMart's shopper base",
        "size": 11, "color": GREY, "font": BODY}])
needed(s, 7.2 + 5.4 * 64 / 100.0 + 0.15, 5.55)
footer(s, 2)
notes(s, "Frame the problem in DMart's language: basket size and repeat rate. The 75% and 64% are "
          "real signals from the study's survey, but flag the sample honestly \u2014 ~50 responses "
          "from college/family, not a DMart panel. The loose-grain 64% is context, not our hero; "
          "the seating gap is. Bridge: tired shoppers want to 'bill and exit,' so the wait directly "
          "threatens unplanned purchases and loyalty.")

# SLIDE 3
s = new_slide(prs)
header(s, "THE INSIGHT  \u00b7  03",
       "The need to rest peaks at the billing counter \u2014 so the rest belongs there, not in a corner.")
journey(s, 2.85,
        [("ARRIVE", "find trolley"), ("BROWSE", "collect items"),
         ("PAUSE", "look for a seat"), ("CHECK OUT", "wait in line"), ("LEAVE", "bill & exit")],
        hi={2, 3})
rect(s, 0.7, 4.25, 11.93, 1.55, fill=SAGE_PALE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
block(s, 1.1, 4.5, 11.2, 1.1, [
    {"text": "THE AHA", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 4},
    {"text": "The study's journey map shows the resting problem surfaces only after ~1 hour \u2014 right at billing, when the shopper is nearly done. A fixed seating zone makes crowding worse. So we bring the rest to the shopper, on the trolley, at the queue.",
     "size": 15.5, "color": INK, "font": BODY, "line": 1.16},
])
block(s, 0.7, 6.05, 11.9, 0.7,
      [{"text": "Implication for DMart: the fix must live in the counter footprint \u2014 zero new floor space, zero extra queue.",
        "size": 13, "color": GREY, "font": BODY}])
footer(s, 3)
notes(s, "This is the pivot slide. The research already did the hard work: it located the pain "
          "precisely \u2014 ~1 hour in, at billing. That tells us a lounge or benches would backfire "
          "(more crowding). The insight is to attach the rest to the asset the shopper already has: "
          "the trolley. Keep the counter footprint unchanged. Transition into how we chose this.")

# SLIDE 4
s = new_slide(prs)
header(s, "IDEATION  \u00b7  04",
       "Four problem angles, four design rules \u2014 one resting concept cleared them all.")
probs = [("NO SEATING", "no rest point", True),
         ("LOOSE-GRAIN BAGS", "misplaced / confusion", False),
         ("TROLLEY PROBLEMS", "wheels / manoeuvre", False),
         ("HIGH SHELVES", "hard to reach", False)]
bw, bh, gx, gy = 2.85, 1.55, 0.30, 0.30
x0, y0 = 0.7, 2.75
for i, (t, sub, chosen) in enumerate(probs):
    r, c = divmod(i, 2)
    x = x0 + c * (bw + gx)
    y = y0 + r * (bh + gy)
    fill = DARK if chosen else WHITE
    line = DARK if chosen else LINE
    rect(s, x, y, bw, bh, fill=fill, line=line, line_w=1.25, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, x + 0.15, y + 0.22, bw - 0.3, bh - 0.4, [
        {"text": t, "size": 13.5, "bold": True, "color": WHITE if chosen else INK, "font": HEAD},
        {"text": sub, "size": 11, "color": WHITE if chosen else GREY, "font": BODY, "space_before": 3},
    ])
    if chosen:
        block(s, x + bw - 1.15, y + 0.12, 1.0, 0.35,
              [{"text": "CHOSEN", "size": 9, "bold": True, "color": WHITE, "font": HEAD, "align": PP_ALIGN.RIGHT}])
block(s, 7.05, 2.75, 5.55, 0.6,
      [{"text": "WHY CUSTOMERS, NOT STAFF", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
block(s, 7.05, 3.12, 5.55, 1.0,
      [{"text": "The team could not get primary data from staff or retail chains, so it focused on the shopper \u2014 and on 'no seating' as the single main problem.",
        "size": 13.5, "color": INK, "font": BODY, "line": 1.14}])
block(s, 7.05, 4.25, 5.55, 0.4,
      [{"text": "FOUR RULES THE WINNER HAD TO PASS", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
crit = ["Independency \u2014 no help, just fold & rest",
        "Body support \u2014 back/arm, fits different users",
        "No mess \u2014 no queue / aisle obstruction",
        "Durable & daily \u2014 cleanable, long-lasting"]
cy = 4.65
for c in crit:
    rect(s, 7.05, cy, 0.34, 0.34, fill=SAGE, shape=MSO_SHAPE.OVAL)
    block(s, 7.05, cy + 0.02, 0.34, 0.3, [{"text": "\u2713", "size": 13, "bold": True,
          "color": WHITE, "font": HEAD, "align": PP_ALIGN.CENTER}])
    block(s, 7.5, cy - 0.02, 5.1, 0.4, [{"text": c, "size": 12.5, "color": INK, "font": BODY}])
    cy += 0.5
footer(s, 4)
notes(s, "Show the discipline: they did not jump to a solution. Four candidate problems, then a hard "
          "constraint \u2014 staff data was unavailable, so they bet on the shopper. Among shopper "
          "problems, 'no seating' won. Then four design rules acted as a filter; only a trolley-"
          "mounted rest could pass all four. That is the concept we are pitching.")

# SLIDE 5
s = new_slide(prs)
header(s, "PROTOTYPE  \u00b7  05",
       "From cardboard mock-up to a brake-locked trolley seat \u2014 we solved the one flaw that could sink it.")
block(s, 0.7, 2.7, 6.0, 0.4, [{"text": "HOW WE BUILT IT", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
path = ["Mini mock-up", "Blender 3D models", "Mechanism POCs", "Prototype envelope"]
px = 0.7
for i, p in enumerate(path):
    rect(s, px, 3.15, 1.35, 0.62, fill=(DARK if i == 3 else SAGE_PALE), line=SAGE, line_w=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, px, 3.27, 1.35, 0.4, [{"text": p, "size": 10.5, "bold": True,
          "color": (WHITE if i == 3 else DARK), "font": HEAD, "align": PP_ALIGN.CENTER}])
    if i < 3:
        block(s, px + 1.35, 3.27, 0.28, 0.4, [{"text": "\u2192", "size": 16, "bold": True,
              "color": SAGE, "font": HEAD, "align": PP_ALIGN.CENTER}])
    px += 1.63
block(s, 0.7, 4.15, 6.1, 2.4, [
    {"text": "MATERIALS", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 3},
    {"text": "Cushion: coconut fibre / firm high-density PU foam. Cover: silicon sheet or TPU. Paddle: recyclable plastic.",
     "size": 13, "color": INK, "font": BODY, "line": 1.12, "space_after": 10, "bullet": True},
    {"text": "SIZE & FINISH", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 3},
    {"text": "Seat 1.5\u20132\" high, 10\u201312\" long, 8\u20139\" wide. Rounded edges, firm, wipe-clean surface.",
     "size": 13, "color": INK, "font": BODY, "line": 1.12, "space_after": 10, "bullet": True},
    {"text": "WHAT WE LEARNED", "size": 12, "color": CLAY, "bold": True, "font": HEAD, "space_after": 3},
    {"text": "Seated weight pushed the cart forward \u2192 solved with a wheel-lock brake. Materials chosen weather / heat / water-proof and easy-clean.",
     "size": 13, "color": INK, "font": BODY, "line": 1.12, "bullet": True},
])
cx, cy = 8.4, 3.05
rect(s, cx, cy, 3.5, 2.55, fill=WHITE, line=DARK, line_w=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, cx + 0.35, cy + 0.35, 2.8, 1.85, fill=SAGE_PALE, line=SAGE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
block(s, cx, cy + 0.95, 3.5, 0.5, [{"text": "SEAT PAD", "size": 12, "bold": True,
      "color": DARK, "font": HEAD, "align": PP_ALIGN.CENTER}])
block(s, cx + 3.55, cy + 1.05, 1.3, 0.5, [{"text": "W 8\u20139\"", "size": 11, "bold": True, "color": INK, "font": HEAD}])
rect(s, cx + 3.45, cy + 1.25, 0.07, 0.07, fill=INK)
block(s, cx - 1.55, cy + 1.05, 1.4, 0.5, [{"text": "L 10\u201312\"", "size": 11, "bold": True, "color": INK, "font": HEAD, "align": PP_ALIGN.RIGHT}])
block(s, cx + 0.6, cy - 0.55, 2.3, 0.4, [{"text": "H 1.5\u20132\"", "size": 11, "bold": True, "color": INK, "font": HEAD, "align": PP_ALIGN.CENTER}])
block(s, 8.4, 5.75, 3.7, 0.7, [{"text": "Prototype envelope \u2014 built for real bodies",
      "size": 11.5, "color": GREY, "font": BODY, "italic": True, "align": PP_ALIGN.CENTER}])
footer(s, 5)
notes(s, "Proof it is real, not a sketch. The team went mock-up \u2192 Blender \u2192 mechanism POC "
          "\u2192 a built prototype envelope with concrete materials and dimensions. The critical "
          "learning: sitting shifts the cart's weight forward, so they engineered a wheel-lock "
          "brake. That is the difference between a student idea and a pilot-ready product. Next: "
          "what it actually is.")

# SLIDE 6
s = new_slide(prs)
header(s, "THE PRODUCT  \u00b7  06",
       "A trolley that doubles as a seat \u2014 rest at the counter, then keep shopping.")
rect(s, 0.7, 2.7, 11.93, 1.0, fill=SAGE_PALE, line=DARK, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
block(s, 1.05, 2.86, 11.3, 0.7, [
    {"text": "IN ONE SENTENCE", "size": 11, "color": CLAY, "bold": True, "font": HEAD, "space_after": 2},
    {"text": "A foldable resting seat that clips onto the DMart trolley so shoppers sit while queuing at billing \u2014 with a wheel-lock brake for safety.",
     "size": 15.5, "bold": True, "color": INK, "font": BODY, "line": 1.08},
])
steps = [("1", "FOLD OUT", "Unclip the paddle; it folds flat from the trolley frame."),
         ("2", "SIT + BRAKE", "Shopper rests; the brake locks the wheels \u2014 no drift."),
         ("3", "FOLD AWAY", "Clip it back; roll straight to the billing counter & exit.")]
sw, sg = 3.75, 0.5
sx = 0.7
for i, (num, t, d) in enumerate(steps):
    x = sx + i * (sw + sg)
    rect(s, x, 4.05, sw, 2.25, fill=WHITE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    oval(s, x + 0.3, 4.35, 0.85, 0.85, fill=DARK)
    block(s, x + 0.3, 4.5, 0.85, 0.6, [{"text": num, "size": 30, "bold": True, "color": WHITE,
          "font": HEAD, "align": PP_ALIGN.CENTER}])
    block(s, x + 1.3, 4.42, sw - 1.5, 0.6, [{"text": t, "size": 15, "bold": True, "color": DARK, "font": HEAD}])
    block(s, x + 0.3, 5.45, sw - 0.6, 0.8, [{"text": d, "size": 12.5, "color": GREY, "font": BODY, "line": 1.12}])
    if i < 2:
        block(s, x + sw + 0.02, 4.9, sg, 0.5, [{"text": "\u2192", "size": 22, "bold": True,
              "color": SAGE, "font": HEAD, "align": PP_ALIGN.CENTER}])
footer(s, 6)
notes(s, "Say it in one sentence and move on \u2014 the Head of Marketing does not need mechanics, "
          "just the mental model. Three steps, all self-service: fold, sit (brake holds the cart), "
          "fold away. No staff, no new floor space, no queue change. This is the product we are "
          "asking them to pilot.")

# SLIDE 7
s = new_slide(prs)
header(s, "KEY FUNCTIONS  \u00b7  07",
       "Five built-in features that turn a wait into a branded DMart moment.")
feats = [
    ("FOLD & REST", "No help needed \u2014 pure self-service.", "Frees floor helpers; on-brand 'everyday ease'."),
    ("BODY SUPPORT", "Back & arm support; options for different users.", "Serves elderly & pregnant \u2014 DMart's loyal base."),
    ("WHEEL-LOCK BRAKE", "Counterweight problem, solved.", "Safe in tight counter space; no aisle block."),
    ("WIPE-CLEAN BUILD", "Weather / heat / water-proof, durable.", "Low upkeep; meets 'cleanliness is a must'."),
    ("CLEAR OWNERSHIP", "Recyclable-plastic paddle, hygienic.", "Sustainable and easy to maintain daily."),
]
cw, cg = 2.30, 0.20
x0 = 0.7
for i, (t, d, b) in enumerate(feats):
    x = x0 + i * (cw + cg)
    rect(s, x, 2.85, cw, 3.7, fill=WHITE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, 2.85, cw, 0.12, DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    oval(s, x + cw/2 - 0.42, 3.15, 0.84, 0.84, fill=SAGE)
    block(s, x + cw/2 - 0.42, 3.32, 0.84, 0.5, [{"text": f"{i+1:02d}", "size": 22, "bold": True,
          "color": DARK, "font": HEAD, "align": PP_ALIGN.CENTER}])
    block(s, x + 0.18, 4.15, cw - 0.36, 0.7, [{"text": t, "size": 13.5, "bold": True,
          "color": INK, "font": HEAD, "align": PP_ALIGN.CENTER, "line": 0.98}])
    block(s, x + 0.18, 4.85, cw - 0.36, 0.9, [{"text": d, "size": 11, "color": GREY, "font": BODY,
          "align": PP_ALIGN.CENTER, "line": 1.08}])
    rect(s, x + 0.18, 5.75, cw - 0.36, 0.012, SAGE)
    block(s, x + 0.18, 5.85, cw - 0.36, 0.7, [{"text": "DMart benefit", "size": 9.5, "bold": True,
          "color": CLAY, "font": HEAD, "align": PP_ALIGN.CENTER, "space_after": 2}])
    block(s, x + 0.18, 6.12, cw - 0.36, 0.5, [{"text": b, "size": 10.5, "color": INK, "font": BODY,
          "align": PP_ALIGN.CENTER, "line": 1.05}])
footer(s, 7)
notes(s, "Tie every feature back to a DMart benefit, not a spec. Fold-and-rest reduces calls on "
          "floor staff; body support serves the older/pregnant base DMart wants to keep; the brake "
          "is the safety answer to the prototype's weight problem; wipe-clean meets the cleanliness "
          "bar the study itself set; clear ownership keeps it hygienic and sustainable. Five reasons "
          "this is a marketing asset, not a cost.")

# SLIDE 8
s = new_slide(prs)
header(s, "DESIGN  \u00b7  08",
       "Built around real bodies and three shopper scenarios \u2014 not a generic bench.")
block(s, 0.7, 2.65, 7.0, 0.35, [{"text": "REIMAGINED JOURNEY", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
journey(s, 3.05,
        [("ARRIVE", "trolley"), ("BROWSE", "collect"), ("PAUSE", "seat appears"),
         ("CHECK OUT", "rest & queue"), ("LEAVE", "fold & exit")],
        hi={2, 3})
block(s, 8.3, 2.65, 4.3, 0.35, [{"text": "THREE SCENARIOS SERVED", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
scen = ["Independent shopper \u2014 visible rest + standing support",
        "Family companion \u2014 sit without losing the group",
        "Temporary fatigue \u2014 a short, judgement-free pause"]
sy = 3.05
for sc in scen:
    rect(s, 8.3, sy, 4.3, 0.78, fill=SAGE_PALE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, 8.5, sy + 0.12, 4.0, 0.55, [{"text": sc, "size": 12, "color": INK, "font": BODY, "line": 1.08}])
    sy += 0.92
block(s, 0.7, 5.35, 7.0, 0.35, [{"text": "DESIGN PRINCIPLES", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
princ = ["Independency", "Body support", "No mess", "Durable & daily"]
px = 0.7
for p in princ:
    rect(s, px, 5.75, 1.62, 0.6, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, px, 5.86, 1.62, 0.4, [{"text": p, "size": 11.5, "bold": True, "color": WHITE, "font": HEAD, "align": PP_ALIGN.CENTER}])
    px += 1.74
block(s, 8.3, 5.95, 4.3, 0.6, [{"text": "Mantra from the study: \u201cbuild for real bodies.\u201d",
      "size": 11.5, "color": GREY, "font": BODY, "italic": True, "align": PP_ALIGN.CENTER}])
footer(s, 8)
notes(s, "Show this is human-centred, not decoration. The journey now has a PAUSE that appears at "
          "the counter; three real scenarios from the study are covered; and the four principles are "
          "the same rules that filtered the concept. The takeaway for marketing: this fits DMart's "
          "bodies and DMart's floor \u2014 not a showroom fantasy.")

# SLIDE 9
s = new_slide(prs)
header(s, "MEASUREMENT  \u00b7  09",
       "We'll track footfall, basket size, repeat rate & campaign ROI \u2014 but baselines and targets need DMart's data.",
       hcolor=INK)
rows = [("KPI", "Baseline", "Target", "How we measure"),
        ("Footfall (pilot stores)", "[NEEDED]", "[NEEDED]", "Store traffic vs control"),
        ("Basket size", "[NEEDED]", "[NEEDED]", "Avg ticket at billing"),
        ("Repeat rate", "[NEEDED]", "[NEEDED]", "Return visits / loyalty"),
        ("Campaign ROI", "[NEEDED]", "[NEEDED]", "Spend vs incremental sales")]
tbl = s.shapes.add_table(len(rows), 4, Inches(0.7), Inches(2.75), Inches(8.7), Inches(3.4)).table
tbl.columns[0].width = Inches(2.7)
tbl.columns[1].width = Inches(1.7)
tbl.columns[2].width = Inches(1.7)
tbl.columns[3].width = Inches(2.6)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.margin_left = Inches(0.08); cell.margin_top = Inches(0.04)
        cell.margin_right = Inches(0.06); cell.margin_bottom = Inches(0.04)
        tf = cell.text_frame; tf.word_wrap = True
        para = tf.paragraphs[0]
        run = para.add_run(); run.text = val
        is_head = (r == 0)
        is_need = (val == "[NEEDED]")
        run.font.size = Pt(11.5 if not is_head else 12)
        run.font.bold = is_head or is_need
        run.font.name = HEAD if is_head else BODY
        run.font.color.rgb = WHITE if is_head else (CLAY if is_need else INK)
        if is_head:
            cell.fill.solid(); cell.fill.fore_color.rgb = DARK
        elif is_need:
            cell.fill.solid(); cell.fill.fore_color.rgb = CLAY_PALE
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb = (SAGE_PALE if r % 2 else WHITE)
block(s, 9.65, 2.75, 2.95, 0.4, [{"text": "MEASUREMENT METHOD", "size": 12, "color": CLAY, "bold": True, "font": HEAD}])
block(s, 9.65, 3.15, 2.95, 2.6, [
    {"text": "Reuse the study's own method: survey-based primary research (the same approach that produced the 75% / 64% signals).",
     "size": 12.5, "color": INK, "font": BODY, "line": 1.14, "space_after": 10, "bullet": True},
    {"text": "Compare pilot stores vs a control store.", "size": 12.5, "color": INK, "font": BODY, "line": 1.14, "bullet": True},
])
needed(s, 9.65, 5.55)
block(s, 9.65 + 1.35, 5.57, 1.7, 0.4, [{"text": "sample & cadence", "size": 10, "color": GREY, "font": BODY}])
footer(s, 9)
notes(s, "Be upfront: the source deck has zero marketing KPIs, so every baseline and target is "
          "[NEEDED] and must come from DMart's own data. What we CAN promise is the method \u2014 the "
          "same survey-based primary research that surfaced the 75%/64% signals, run pilot-vs-control. "
          "Ask the marketing head to supply the baselines so we can set real targets together.")

# SLIDE 10
s = new_slide(prs)
header(s, "THE ASK & NEXT STEPS  \u00b7  10",
       "Pilot the trolley-seat in rush-hour billing zones \u2014 scope, cost & timeline to set with DMart.")
ask = [
    ("PILOT SCOPE", "Rush-hour billing queues in high-footfall stores; focus on older / pregnant-heavy catchment.",
     "store count & geographies"),
    ("TIMELINE & INVESTMENT", "Build + field the prototype; measure at midpoint and close.",
     "weeks, capex & opex"),
    ("EXPECTED OUTCOME", "Validate the 75% rest-pain signal in DMart's stores; test the loyalty / queue-ease hypothesis.",
     "quantified targets"),
]
aw, ag = 3.75, 0.5
ax = 0.7
for i, (t, d, need) in enumerate(ask):
    x = ax + i * (aw + ag)
    rect(s, x, 2.85, aw, 3.3, fill=WHITE, line=LINE, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, 2.85, aw, 0.62, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    block(s, x, 2.93, aw, 0.45, [{"text": t, "size": 14, "bold": True, "color": WHITE, "font": HEAD, "align": PP_ALIGN.CENTER}])
    block(s, x + 0.25, 3.7, aw - 0.5, 1.7, [{"text": d, "size": 13, "color": INK, "font": BODY, "line": 1.16}])
    needed(s, x + 0.25, 5.45)
    block(s, x + 0.25 + 1.4, 5.47, aw - 1.6, 0.4, [{"text": need, "size": 10.5, "color": GREY, "font": BODY}])
rect(s, 0.7, 6.45, 11.93, 0.55, fill=DARK, line=DARK, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
block(s, 0.7, 6.5, 11.93, 0.45, [
    {"text": "THE ASK:  approve a low-risk pilot  \u2014  we bring the prototype, DMart brings the stores and the data.",
     "size": 13.5, "bold": True, "color": CREAM, "font": HEAD, "align": PP_ALIGN.CENTER}])
footer(s, 10)
notes(s, "Close with a small, safe ask \u2014 not a national rollout. A rush-hour billing-zone pilot "
          "in a few high-footfall stores, using the prototype we already built. We are explicit that "
          "scope, cost and timeline are [NEEDED] from DMart, because the study gave us the problem "
          "and the product, not the P&L. Offer to co-own the measurement plan from Slide 9.")

OUT = "/home/user/DMart_TrolleyRest_Pitch.pptx"
prs.save(OUT)
print("SAVED:", OUT, "slides:", len(prs.slides._sldIdLst))
