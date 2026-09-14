import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE

BASE = "/Users/nirbhayaggarwal/Documents/Default Project/Nirbhay_Test/wedding-invite"
OUT = os.path.expanduser("~/Downloads/Nirbhay & Riya - Partner Review.pptx")
URL = "https://api.tamarind.sh/d/nirbhay-riya-wedding-3fc1ff/"

MAROON = RGBColor(0x9D, 0x2B, 0x23)
INK = RGBColor(0x37, 0x31, 0x2A)
GRAY = RGBColor(0x6D, 0x64, 0x55)
LIGHT_GRAY = RGBColor(0xA9, 0xA1, 0x92)
CREAM = RGBColor(0xFA, 0xF6, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHIP_LINE = RGBColor(0xE0, 0xD8, 0xC8)

TITLE_F = "Georgia"
BODY_F = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
W, H = prs.slide_width, prs.slide_height


def bg(slide, color=WHITE):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False
    return r


def tb(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    first = True
    for spec in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = spec.get("align", align)
        if "space_before" in spec:
            p.space_before = Pt(spec["space_before"])
        if "space_after" in spec:
            p.space_after = Pt(spec["space_after"])
        for txt, f, s, c, b, i in spec["runs"]:
            r = p.add_run()
            r.text = txt
            r.font.name = f
            r.font.size = Pt(s)
            r.font.color.rgb = c
            r.font.bold = b
            r.font.italic = i
    return box


def R(txt, size, color=INK, bold=False, font=BODY_F, italic=False):
    return (txt, font, size, color, bold, italic)


def title_bar(slide, text, page, sub=None):
    tb(slide, Inches(0.55), Inches(0.32), Inches(12.2), Inches(1.2),
       [{"runs": [R(text, 22, INK, True, TITLE_F)]}])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.18), Inches(0.9), Pt(3))
    bar.fill.solid()
    bar.fill.fore_color.rgb = MAROON
    bar.line.fill.background()
    bar.shadow.inherit = False
    if sub:
        tb(slide, Inches(0.55), Inches(1.32), Inches(12.2), Inches(0.4),
           [{"runs": [R(sub, 11, GRAY)]}])
    tb(slide, Inches(0.55), Inches(7.06), Inches(9), Inches(0.3),
       [{"runs": [R("Source: wedding-invite/index.html + live site (link on cover)", 8, LIGHT_GRAY)]}])
    tb(slide, Inches(12.35), Inches(7.06), Inches(0.5), Inches(0.3),
       [{"runs": [R(str(page), 9, GRAY)]}], align=PP_ALIGN.RIGHT)


def chip(slide, x, y, w, h, num, label, num_color=MAROON):
    c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    c.adjustments[0] = 0.10
    c.fill.solid()
    c.fill.fore_color.rgb = WHITE
    c.line.color.rgb = CHIP_LINE
    c.line.width = Pt(0.75)
    c.shadow.inherit = False
    tf = c.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.14)
    tf.margin_top = Inches(0.08)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = num; r.font.name = TITLE_F; r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = num_color
    p2 = tf.add_paragraph()
    r2 = p2.add_run(); r2.text = label; r2.font.name = BODY_F; r2.font.size = Pt(10.5); r2.font.color.rgb = INK
    return c


# ---------------- Slide 1: cover ----------------
s = prs.slides.add_slide(BLANK)
bg(s, CREAM)
frame = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.75), Inches(1.05), Inches(3.7), Inches(5.35))
frame.fill.solid(); frame.fill.fore_color.rgb = WHITE; frame.line.color.rgb = CHIP_LINE; frame.line.width = Pt(1)
frame.shadow.inherit = False
s.shapes.add_picture(f"{BASE}/assets/gallery/p3.jpg", Inches(8.9), Inches(1.2), width=Inches(3.4))
tb(s, Inches(0.7), Inches(1.15), Inches(7.6), Inches(0.5),
   [{"runs": [R("P A R T N E R   R E V I E W   ·   A U G U S T   2 0 2 6", 11, GRAY)]}])
tb(s, Inches(0.7), Inches(1.75), Inches(7.6), Inches(2.6),
   [{"runs": [R("Nirbhay & Riya —", 30, INK, True, TITLE_F)]},
    {"runs": [R("the wedding invitation website is live and ready for your review", 30, INK, True, TITLE_F)], "space_before": 4}])
bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.72), Inches(4.05), Inches(1.2), Pt(3.2))
bar.fill.solid(); bar.fill.fore_color.rgb = MAROON; bar.line.fill.background(); bar.shadow.inherit = False
tb(s, Inches(0.7), Inches(4.35), Inches(7.6), Inches(1.6),
   [{"runs": [R("One page · 11 sections · motion-luxury design · all confirmed details included", 14, GRAY)], "space_after": 8},
    {"runs": [R(URL, 13, MAROON, True)]},
    {"runs": [R("Works on phone and laptop · December 5–7, 2026 · Noura, The Palace, Chandigarh", 11, LIGHT_GRAY)], "space_before": 8}])

# ---------------- Slide 2: at a glance ----------------
s = prs.slides.add_slide(BLANK)
bg(s)
title_bar(s, "The invitation is one live page with 11 sections covering story, schedule, photos and RSVP", 2)
sections = [
    ("01", "Hero — cordially invited + couple photo"),
    ("02", "“These kids are getting married!” — childhood polaroids"),
    ("03", "Families — parents’ names, both sides"),
    ("04", "Date banner — Sunday, 07 December 2026"),
    ("05", "How we met — Ahmedabad ↔ Bodhgaya journey"),
    ("06", "Countdown — live, to Dec 5"),
    ("07", "Ceremonial details — 3 days of events"),
    ("08", "Gallery — curated couple photos"),
    ("09", "“Forever” — B&W torn photo strip"),
    ("10", "Location — Noura, The Palace + map"),
    ("11", "RSVP — Aggarwal Family · #NiRi"),
]
x0, y0, cw, ch, gx, gy = Inches(0.55), Inches(1.7), Inches(3.0), Inches(1.12), Inches(0.25), Inches(0.22)
for i, (n, lbl) in enumerate(sections):
    x = x0 + (i % 3) * (cw + gx)
    y = y0 + (i // 3) * (ch + gy)
    chip(s, x, y, cw, ch, n, lbl)
rail = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.35), Inches(1.7), Inches(2.45), Inches(5.05))
rail.adjustments[0] = 0.05
rail.fill.solid(); rail.fill.fore_color.rgb = CREAM; rail.line.color.rgb = CHIP_LINE; rail.line.width = Pt(0.75); rail.shadow.inherit = False
tb(s, Inches(10.6), Inches(1.95), Inches(2.0), Inches(4.6),
   [{"runs": [R("Status", 11, MAROON, True, TITLE_F)], "space_after": 2},
    {"runs": [R("Live since Aug 31, 2026", 10, INK)], "space_after": 12},
    {"runs": [R("Devices", 11, MAROON, True, TITLE_F)], "space_after": 2},
    {"runs": [R("Responsive — phone + laptop", 10, INK)], "space_after": 12},
    {"runs": [R("Stack", 11, MAROON, True, TITLE_F)], "space_after": 2},
    {"runs": [R("Single static page, GSAP motion, hosted URL", 10, INK)], "space_after": 12},
    {"runs": [R("Photos", 11, MAROON, True, TITLE_F)], "space_after": 2},
    {"runs": [R("6 of 7 couple shots used + 2 childhood", 10, INK)]}])

# ---------------- Slide 3: design ----------------
s = prs.slides.add_slide(BLANK)
bg(s)
title_bar(s, "The look follows the approved reference — airy cream with floral line art — plus 6 motion-luxury layers", 3)
col_l = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.7), Inches(6.0), Inches(4.9))
col_l.adjustments[0] = 0.04
col_l.fill.solid(); col_l.fill.fore_color.rgb = CREAM; col_l.line.color.rgb = CHIP_LINE; col_l.line.width = Pt(0.75); col_l.shadow.inherit = False
tb(s, Inches(0.85), Inches(1.95), Inches(5.5), Inches(4.4),
   [{"runs": [R("Reference style — approved", 14, MAROON, True, TITLE_F)], "space_after": 8}] +
   [{"runs": [R("·  " + t, 11.5, INK)], "space_after": 6} for t in [
       "Airy cream canvas, floral line-art corners",
       "Torn-paper edges on photo sections",
       "Serif caps (Cormorant) + readable script (Alex Brush)",
       "Maroon + gold accents, gold venue pills",
       "Polaroid photos with tilted frames and hearts",
   ]])
col_r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.7), Inches(6.0), Inches(4.9))
col_r.adjustments[0] = 0.04
col_r.fill.solid(); col_r.fill.fore_color.rgb = WHITE; col_r.line.color.rgb = MAROON; col_r.line.width = Pt(1.2); col_r.shadow.inherit = False
tb(s, Inches(7.1), Inches(1.95), Inches(5.5), Inches(4.4),
   [{"runs": [R("Motion-luxury layers — new", 14, MAROON, True, TITLE_F)], "space_after": 8}] +
   [{"runs": [R(f"{i+1}.  " + t, 11.5, INK)], "space_after": 5} for i, t in enumerate([
       "Gold-foil shimmer on the couple’s names",
       "Letter-by-letter hero name reveal",
       "Parallax hero + B&W strip on scroll",
       "Ken Burns slow zoom on the hero photo",
       "Journey, gallery and details stagger in on scroll",
       "Film grain overlay + dark-luxe footer (graceful fallback if animations can’t load)",
   ])])

# ---------------- Slide 4: story ----------------
s = prs.slides.add_slide(BLANK)
bg(s)
title_bar(s, "“Miles apart, meant to be” — he was in Ahmedabad, she was in Bodhgaya, forever begins in Chandigarh", 4)
nodes = [("Ahmedabad", "he was here", Inches(2.1)), ("Chandigarh", "where forever begins", Inches(6.15)), ("Bodhgaya", "she was here", Inches(10.2))]
for i, (city, cap, cx) in enumerate(nodes):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, cx, Inches(1.95), Inches(0.34), Inches(0.34))
    c.fill.solid(); c.fill.fore_color.rgb = MAROON if city == "Chandigarh" else WHITE
    c.line.color.rgb = MAROON; c.line.width = Pt(1.5); c.shadow.inherit = False
    tb(s, cx - Inches(1.0), Inches(2.4), Inches(2.35), Inches(0.75),
       [{"runs": [R(city, 15, INK, True, TITLE_F)], "align": PP_ALIGN.CENTER},
        {"runs": [R(cap, 10, GRAY)], "align": PP_ALIGN.CENTER}])
for x1, x2 in [(Inches(2.44), Inches(6.15)), (Inches(6.49), Inches(10.2))]:
    conn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, Inches(2.12), x2, Inches(2.12))
    conn.line.color.rgb = LIGHT_GRAY; conn.line.width = Pt(1.5); conn.line.dash_style = MSO_LINE.DASH
    conn.shadow.inherit = False
y = Inches(3.6)
from PIL import Image as PILImage
for j, p in enumerate(["p7", "p2", "p6"]):
    fx = Inches(1.15) + j * Inches(3.95)
    fr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, fx, y, Inches(3.1), Inches(2.5))
    fr.fill.solid(); fr.fill.fore_color.rgb = WHITE; fr.line.color.rgb = CHIP_LINE; fr.line.width = Pt(1); fr.shadow.inherit = False
    img = PILImage.open(f"{BASE}/assets/gallery/{p}.jpg")
    ph = Inches(2.34)
    pw = int(ph * img.size[0] / img.size[1])
    s.shapes.add_picture(f"{BASE}/assets/gallery/{p}.jpg", int(fx + (Inches(3.1) - pw) / 2), y + Inches(0.08), height=ph, width=pw)
tb(s, Inches(0.55), Inches(6.35), Inches(12.2), Inches(0.5),
   [{"runs": [R("Curated gallery — 6 of the 7 couple photos are used across the page (hero, gallery, B&W strip)", 11, GRAY)], "align": PP_ALIGN.CENTER}])

# ---------------- Slide 5: details ----------------
s = prs.slides.add_slide(BLANK)
bg(s)
title_bar(s, "All confirmed details are in — 3 days of events, December 5–7, at Noura, The Palace, Chandigarh", 5)
rows = [
    ("Friday · 05 Dec", "Mehendi", "—"),
    ("Saturday · 06 Dec", "Haldi (day)  ·  Ring Ceremony & Sangeet (evening)", "Noura, The Palace"),
    ("Sunday · 07 Dec", "Wedding — sundowner ceremony", "Noura, The Palace"),
]
ty = Inches(1.75)
tb(s, Inches(0.95), ty - Inches(0.34), Inches(2.6), Inches(0.3), [{"runs": [R("DAY", 10, GRAY, True)]}])
tb(s, Inches(3.75), ty - Inches(0.34), Inches(5.9), Inches(0.3), [{"runs": [R("EVENTS", 10, GRAY, True)]}])
tb(s, Inches(9.85), ty - Inches(0.34), Inches(2.7), Inches(0.3), [{"runs": [R("VENUE", 10, GRAY, True)]}])
for i, (day, ev, ven) in enumerate(rows):
    ry = ty + i * Inches(1.0)
    r_ = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), ry, Inches(11.95), Inches(0.88))
    r_.adjustments[0] = 0.12
    r_.fill.solid(); r_.fill.fore_color.rgb = CREAM if i % 2 == 0 else WHITE
    r_.line.color.rgb = CHIP_LINE; r_.line.width = Pt(0.75); r_.shadow.inherit = False
    tb(s, Inches(0.95), ry + Inches(0.24), Inches(2.6), Inches(0.5), [{"runs": [R(day, 12.5, MAROON, True, TITLE_F)]}])
    tb(s, Inches(3.75), ry + Inches(0.24), Inches(5.9), Inches(0.5), [{"runs": [R(ev, 12, INK)]}])
    tb(s, Inches(9.85), ry + Inches(0.24), Inches(2.7), Inches(0.5), [{"runs": [R(ven, 11, GRAY)]}])
tb(s, Inches(0.7), Inches(5.05), Inches(11.95), Inches(1.7),
   [{"runs": [R("Families", 12, MAROON, True, TITLE_F)], "space_after": 3},
    {"runs": [R("Nirbhay — son of Mr. Sanjiv Aggarwal & Mrs. Krishna Aggarwal", 11.5, INK)], "space_after": 2},
    {"runs": [R("Riya — daughter of Mr. Nikesh Gupta & Mrs. Sonia Gupta", 11.5, INK)], "space_after": 10},
    {"runs": [R("RSVP", 12, MAROON, True, TITLE_F), R("   Aggarwal Family        ", 11.5, INK), R("Hashtag", 12, MAROON, True, TITLE_F), R("   #NiRi", 11.5, INK)]}])

# ---------------- Slide 6: open items ----------------
s = prs.slides.add_slide(BLANK)
bg(s)
title_bar(s, "2 open items before the link goes out widely — everything else is final", 6)
items = [
    ("1", "Mehendi timing + venue (Fri, Dec 5)", "The event is listed for Day 1 with no time or venue yet — the 6th–7th are at Noura, The Palace."),
    ("2", "RSVP contact number (optional)", "RSVP currently reads “Aggarwal Family”. A phone number can be added as a tap-to-call pill, like the reference."),
]
for i, (n, head, body) in enumerate(items):
    cx = Inches(0.7) + i * Inches(6.15)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), Inches(5.8), Inches(2.6))
    card.adjustments[0] = 0.05
    card.fill.solid(); card.fill.fore_color.rgb = WHITE; card.line.color.rgb = MAROON; card.line.width = Pt(1.2); card.shadow.inherit = False
    tb(s, cx + Inches(0.3), Inches(2.05), Inches(5.2), Inches(2.2),
       [{"runs": [R(n + ".  ", 20, MAROON, True, TITLE_F), R(head, 15, INK, True, TITLE_F)], "space_after": 8},
        {"runs": [R(body, 11.5, GRAY)]}])
tb(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(1.6),
   [{"runs": [R("Also parked for now: ", 11.5, GRAY), R("the dress-code section (dropped on your call — one line restores it), and extra gallery photos (p5 unused).", 11.5, GRAY)], "space_after": 10},
    {"runs": [R("Next step: ", 12, INK, True), R("send the two details — one redeploy updates the same live link.", 12, INK)]}])

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides._sldIdLst))
