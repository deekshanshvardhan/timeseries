"""
Generates the strategic recommendation deck:
"Louie Voice SDK - Rapido vs Saarthi: Which app should we voice-enable?"

Run:  python3 build_deck.py
Output:  Louie_Voice_SDK_Rapido_vs_Saarthi.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Design system
# ----------------------------------------------------------------------------
NAVY      = RGBColor(0x14, 0x1B, 0x41)   # primary dark
INDIGO    = RGBColor(0x3B, 0x3D, 0xBF)   # Louie accent
TEAL      = RGBColor(0x00, 0xB3, 0x8A)   # "go" / Rapido positive
AMBER     = RGBColor(0xF5, 0xA6, 0x23)   # Rapido brand-ish
SLATE     = RGBColor(0x5B, 0x61, 0x77)   # muted text
LIGHT     = RGBColor(0xF4, 0xF5, 0xFA)   # light panel
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GREY      = RGBColor(0x9A, 0x9F, 0xB3)
RED       = RGBColor(0xD6, 0x4F, 0x4F)   # Saarthi caution
DARKTEXT  = RGBColor(0x22, 0x25, 0x36)

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill=None, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    return sp


def txt(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=6, line_spacing=1.0, wrap=True):
    """runs = list of paragraphs; each paragraph = list of (text, size, color, bold, italic) tuples."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for run_spec in para:
            t, size, color, bold = run_spec[0], run_spec[1], run_spec[2], run_spec[3]
            italic = run_spec[4] if len(run_spec) > 4 else False
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Calibri"
    return tb


def P(text, size, color=DARKTEXT, bold=False, italic=False):
    """Single-run paragraph shortcut."""
    return [(text, size, color, bold, italic)]


def bullet(text, size=15, color=DARKTEXT, bold=False):
    return [("•  ", size, INDIGO, True, False), (text, size, color, bold, False)]


def header(slide, kicker, title, accent=INDIGO):
    """Standard content-slide header band."""
    rect(slide, 0, 0, EMU_W, Inches(1.25), fill=NAVY)
    rect(slide, 0, Inches(1.25), EMU_W, Pt(4), fill=accent)
    txt(slide, Inches(0.6), Inches(0.18), Inches(12), Inches(0.35),
        [P(kicker.upper(), 12, accent, True)], anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(0.6), Inches(0.5), Inches(12.1), Inches(0.7),
        [P(title, 27, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)


_PAGE = [1]  # title slide is page 1 (no footer); footers auto-number from 2


def footer(slide, page=None):
    _PAGE[0] += 1
    txt(slide, Inches(0.6), Inches(7.06), Inches(8), Inches(0.3),
        [P("Louie Voice SDK  |  Strategic Project Selection  |  Confidential", 9, GREY)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(12.0), Inches(7.06), Inches(0.8), Inches(0.3),
        [P(str(_PAGE[0]), 9, GREY, True)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def placeholder(slide, x, y, w, h, label):
    """A dashed box reminding where to paste an app screenshot."""
    box = rect(slide, x, y, w, h, fill=LIGHT, line=GREY, line_w=Pt(1.25))
    box.line.dash_style = None
    ln = box.line._get_or_add_ln()
    d = ln.makeelement(qn('a:prstDash'), {'val': 'dash'})
    ln.append(d)
    txt(slide, x, y + (h - Inches(0.5)) / 2, w, Inches(0.5),
        [P("[ Paste screenshot here ]", 12, SLATE, True),
         P(label, 10, GREY, False)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=2)


def chip(slide, x, y, w, text, fill, fg=WHITE, h=Inches(0.42), size=12):
    c = rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = 0.5
    txt(slide, x, y, w, h, [P(text, size, fg, True)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ----------------------------------------------------------------------------
# 1. TITLE SLIDE
# ----------------------------------------------------------------------------
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, fill=NAVY)
# accent shapes
rect(s, 0, 0, Inches(0.35), EMU_H, fill=INDIGO)
rect(s, Inches(9.1), 0, Inches(4.6), EMU_H, fill=RGBColor(0x1B, 0x23, 0x52))
rect(s, Inches(9.1), 0, Pt(4), EMU_H, fill=TEAL)

txt(s, Inches(0.9), Inches(1.0), Inches(8), Inches(0.4),
    [P("LOUIE VOICE SDK  ·  STRATEGIC RECOMMENDATION", 13, TEAL, True)])
txt(s, Inches(0.9), Inches(1.7), Inches(8.0), Inches(2.4),
    [P("Rapido over Saarthi", 46, WHITE, True),
     P("Which app should we voice-enable first?", 24, GREY, False)],
    space_after=10)
txt(s, Inches(0.9), Inches(3.9), Inches(7.8), Inches(1.6),
    [P("A resource-crunch decision, made on first principles:", 15, WHITE, True),
     bullet("Play to our core differentiator \u2014 end-to-end voice transactions", 15, RGBColor(0xCF,0xD3,0xE6)),
     bullet("Maximise reach \u2014 mass-market, habitual, low-deliberation usage", 15, RGBColor(0xCF,0xD3,0xE6)),
     bullet("Deliver inclusion impact \u2014 a universal everyday need", 15, RGBColor(0xCF,0xD3,0xE6))],
    space_after=8)

# verdict chip on right panel
chip(s, Inches(9.65), Inches(2.7), Inches(3.5), "RECOMMENDATION", INDIGO, size=12, h=Inches(0.45))
txt(s, Inches(9.65), Inches(3.25), Inches(3.5), Inches(1.0),
    [P("GO WITH", 16, GREY, True), P("RAPIDO", 40, TEAL, True)], align=PP_ALIGN.CENTER, space_after=2)
txt(s, Inches(9.65), Inches(4.6), Inches(3.5), Inches(0.6),
    [P("Bike-Taxi, Auto & Cabs", 13, RGBColor(0xCF,0xD3,0xE6), False)], align=PP_ALIGN.CENTER)

txt(s, Inches(0.9), Inches(6.7), Inches(11), Inches(0.4),
    [P("Prepared by:  [Your Name]   |   Role: Business Analyst   |   Date: [DD Mon YYYY]", 12, GREY)])


# ----------------------------------------------------------------------------
# 2. EXECUTIVE SUMMARY
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Executive summary", "The recommendation in one slide", TEAL)

txt(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.9),
    [[("Choose ", 18, DARKTEXT, False), ("Rapido", 18, TEAL, True),
      (". It is the only one of the two apps whose core user journey is an ", 18, DARKTEXT, False),
      ("end-to-end transaction", 18, DARKTEXT, True),
      (" \u2014 exactly the capability that makes Louie unique.", 18, DARKTEXT, False)]],
    line_spacing=1.05)

cards = [
    ("Differentiator fit", "Rapido = book \u2192 confirm \u2192 pay \u2192 ride, a true end-to-end transaction. Saarthi is an awareness / education app \u2014 there is no transaction to complete.", TEAL),
    ("Decision simplicity", "A ride is a low-stakes, habitual command (\u201cBook an auto home\u201d). Investing is high-stakes deliberation that users will not hand to voice.", INDIGO),
    ("Market reach", "Hundreds of millions ride; only a small, mostly urban minority of India invests \u2014 and even fewer invest impulsively.", AMBER),
    ("Inclusion impact", "Mobility is a universal daily need. Voice unlocks it for low-literacy, vernacular and visually-impaired users far more broadly than investing does.", RED),
]
cx = Inches(0.6)
cw = Inches(2.95)
gap = Inches(0.18)
for i, (h, b, col) in enumerate(cards):
    x = cx + i * (cw + gap)
    rect(s, x, Inches(2.7), cw, Inches(3.55), fill=LIGHT)
    rect(s, x, Inches(2.7), cw, Pt(6), fill=col)
    txt(s, x + Inches(0.2), Inches(2.95), cw - Inches(0.4), Inches(0.7),
        [P(h, 16, col, True)])
    txt(s, x + Inches(0.2), Inches(3.7), cw - Inches(0.4), Inches(2.4),
        [P(b, 13, DARKTEXT)], line_spacing=1.05)

footer(s, 2)


# ----------------------------------------------------------------------------
# 3. THE STRATEGIC QUESTION
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Context", "The decision we have been asked to make", INDIGO)

txt(s, Inches(0.6), Inches(1.6), Inches(12), Inches(0.9),
    [[("Two apps have approached us to voice-enable their product. Due to ongoing commitments, we can onboard ", 16, DARKTEXT, False),
      ("only one", 16, DARKTEXT, True), (".", 16, DARKTEXT, False)]], line_spacing=1.05)

# two option boxes
def option_box(x, name, tag, lines, col):
    rect(s, x, Inches(2.6), Inches(5.9), Inches(3.05), fill=WHITE, line=col, line_w=Pt(1.5))
    rect(s, x, Inches(2.6), Inches(5.9), Inches(0.75), fill=col)
    txt(s, x + Inches(0.25), Inches(2.6), Inches(5.4), Inches(0.75),
        [P(name, 20, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.25), Inches(3.5), Inches(5.4), Inches(0.4),
        [P(tag, 13, col, True)])
    txt(s, x + Inches(0.25), Inches(3.95), Inches(5.4), Inches(1.6),
        [bullet(l, 13) for l in lines], space_after=6)

option_box(Inches(0.6), "Saarthi", "Investment awareness app by SEBI",
           ["Purpose: educate & spread investor awareness",
            "Content-led: articles, videos, quizzes, info",
            "Audience: existing / aspiring retail investors"], INDIGO)
option_box(Inches(6.85), "Rapido", "Bike-Taxi, Auto & Cabs",
           ["Purpose: book and complete a ride",
            "Transaction-led: search \u2192 book \u2192 pay \u2192 travel",
            "Audience: mass-market daily commuters"], AMBER)

txt(s, Inches(0.6), Inches(5.85), Inches(12), Inches(0.8),
    [[("The question: ", 15, NAVY, True),
      ("Which single app best lets Louie leverage its strengths and core differentiators?", 15, DARKTEXT, False)]])
footer(s, 3)


# ----------------------------------------------------------------------------
# 4. CEO GUIDING PRINCIPLE
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Our decision rule", "The CEO's guiding principle", TEAL)

rect(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(1.7), fill=LIGHT)
rect(s, Inches(0.6), Inches(1.7), Pt(6), Inches(1.7), fill=TEAL)
txt(s, Inches(1.0), Inches(1.95), Inches(11.4), Inches(1.3),
    [[("\u201cWhenever we are in a resource crunch, always select the project that lets us ", 20, DARKTEXT, True),
      ("leverage our strengths and our core differentiators", 20, TEAL, True),
      (".\u201d", 20, DARKTEXT, True)]], line_spacing=1.1, anchor=MSO_ANCHOR.MIDDLE)

txt(s, Inches(0.6), Inches(3.8), Inches(12), Inches(0.5),
    [P("So the whole decision reduces to one test:", 17, NAVY, True)])
rect(s, Inches(0.6), Inches(4.4), Inches(12.1), Inches(1.5), fill=NAVY)
txt(s, Inches(1.0), Inches(4.4), Inches(11.4), Inches(1.5),
    [[("Which app's core journey is an ", 22, WHITE, False),
      ("end-to-end transaction in the user's own language", 22, TEAL, True),
      (" \u2014 our single biggest differentiator?", 22, WHITE, False)]],
    anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
footer(s, 4)


# ----------------------------------------------------------------------------
# 5. LOUIE DIFFERENTIATORS
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "What makes us win", "Louie's core differentiators", INDIGO)

diffs = [
    ("End-to-end transactions", "Not just answering questions \u2014 Louie completes the whole task by voice, from intent to confirmation.", TEAL),
    ("Multilingual by default", "Delivers a voice interface in many Indian languages, even when the app itself is English-only.", INDIGO),
    ("No global direct competitor", "Google Assistant, Alexa and Siri trigger steps but do not complete end-to-end transactions.", AMBER),
    ("Sector-agnostic plug-in", "One SDK that voice-enables any mobile app on Android & iOS.", RED),
]
for i, (h, b, col) in enumerate(diffs):
    x = Inches(0.6) + (i % 2) * Inches(6.25)
    y = Inches(1.7) + (i // 2) * Inches(2.45)
    rect(s, x, y, Inches(5.95), Inches(2.2), fill=LIGHT)
    rect(s, x, y, Inches(0.14), Inches(2.2), fill=col)
    txt(s, x + Inches(0.35), y + Inches(0.25), Inches(5.4), Inches(0.6),
        [P(h, 18, col, True)])
    txt(s, x + Inches(0.35), y + Inches(0.95), Inches(5.4), Inches(1.1),
        [P(b, 14, DARKTEXT)], line_spacing=1.05)

txt(s, Inches(0.6), Inches(6.7), Inches(12), Inches(0.4),
    [[("Implication: ", 13, NAVY, True),
      ("our value is highest where there is a real transaction to complete by voice.", 13, SLATE, False)]])
footer(s, 5)


# ----------------------------------------------------------------------------
# 6. SAARTHI SNAPSHOT (screenshots)
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "App snapshot 1 of 2", "Saarthi \u2014 investment awareness (SEBI)", INDIGO)
placeholder(s, Inches(0.6), Inches(1.6), Inches(2.0), Inches(4.0), "Home / learn feed")
placeholder(s, Inches(2.75), Inches(1.6), Inches(2.0), Inches(4.0), "Article / video")
placeholder(s, Inches(4.9), Inches(1.6), Inches(2.0), Inches(4.0), "Quiz / tools")
rect(s, Inches(7.3), Inches(1.6), Inches(5.4), Inches(4.0), fill=LIGHT)
rect(s, Inches(7.3), Inches(1.6), Inches(5.4), Pt(6), fill=INDIGO)
txt(s, Inches(7.6), Inches(1.85), Inches(4.9), Inches(3.6),
    [P("What we observe", 17, INDIGO, True),
     bullet("Primary value is information & awareness, not a transaction", 14),
     bullet("Journeys end in \u201cread / watch / learn\u201d, not \u201cbuy / confirm\u201d", 14),
     bullet("No money actually moves inside the app", 14),
     bullet("Content is the product \u2014 best consumed visually, at the user's own pace", 14),
     bullet("Little for an end-to-end voice flow to complete", 14, RED, True)],
    space_after=8)
txt(s, Inches(0.6), Inches(5.75), Inches(12), Inches(0.5),
    [[("Note: ", 12, SLATE, True),
      ("replace placeholders with real screenshots from the Play Store / App Store build.", 12, SLATE, False)]])
footer(s, 6)


# ----------------------------------------------------------------------------
# 7. RAPIDO SNAPSHOT (screenshots)
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "App snapshot 2 of 2", "Rapido \u2014 Bike-Taxi, Auto & Cabs", AMBER)
placeholder(s, Inches(0.6), Inches(1.6), Inches(2.0), Inches(4.0), "Enter destination")
placeholder(s, Inches(2.75), Inches(1.6), Inches(2.0), Inches(4.0), "Pick ride & fare")
placeholder(s, Inches(4.9), Inches(1.6), Inches(2.0), Inches(4.0), "Confirm & pay")
rect(s, Inches(7.3), Inches(1.6), Inches(5.4), Inches(4.0), fill=LIGHT)
rect(s, Inches(7.3), Inches(1.6), Inches(5.4), Pt(6), fill=AMBER)
txt(s, Inches(7.6), Inches(1.85), Inches(4.9), Inches(3.6),
    [P("What we observe", 17, AMBER, True),
     bullet("A clear, repeatable transaction: search \u2192 select \u2192 pay \u2192 ride", 14),
     bullet("High-frequency, habitual usage \u2014 done many times a week", 14),
     bullet("Few decision variables: where to, which vehicle, confirm", 14),
     bullet("Often used on the move, hands-busy \u2014 voice is genuinely useful", 14),
     bullet("A perfect end-to-end flow for Louie to own", 14, TEAL, True)],
    space_after=8)
txt(s, Inches(0.6), Inches(5.75), Inches(12), Inches(0.5),
    [[("Note: ", 12, SLATE, True),
      ("replace placeholders with real screenshots of the booking flow.", 12, SLATE, False)]])
footer(s, 7)


# ----------------------------------------------------------------------------
# 8. EVALUATION FRAMEWORK
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "How we judged", "Our evaluation framework", TEAL)
txt(s, Inches(0.6), Inches(1.55), Inches(12), Inches(0.5),
    [P("Five criteria, each tied directly to how Louie creates and captures value.", 15, SLATE)])
crit = [
    ("1", "Transaction fit", "Is the core journey an end-to-end transaction we can complete by voice?"),
    ("2", "Decision simplicity", "Can a user comfortably delegate the decision to voice without hesitation?"),
    ("3", "Market reach (TAM)", "How many people perform this action, and how often?"),
    ("4", "Inclusion impact", "Does voice unlock a genuine, universal need for under-served users?"),
    ("5", "Multilingual relevance", "Does a vernacular voice layer meaningfully widen the audience?"),
]
y = Inches(2.25)
for num, h, b in crit:
    rect(s, Inches(0.6), y, Inches(12.1), Inches(0.78), fill=LIGHT)
    c = rect(s, Inches(0.75), y + Inches(0.12), Inches(0.55), Inches(0.55), fill=INDIGO, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.75), y + Inches(0.12), Inches(0.55), Inches(0.55),
        [P(num, 18, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.55), y, Inches(3.4), Inches(0.78),
        [P(h, 16, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(5.0), y, Inches(7.5), Inches(0.78),
        [P(b, 13.5, DARKTEXT)], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.86)
footer(s, 8)


# ----------------------------------------------------------------------------
# 9-13. REASON SLIDES
# ----------------------------------------------------------------------------
def reason_slide(idx, kicker, title, lead, rapido_pts, saarthi_pts, takeaway, page):
    s = add_slide()
    header(s, kicker, title, TEAL)
    txt(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.8),
        [P(lead, 16, DARKTEXT)], line_spacing=1.05)
    # Rapido column
    rect(s, Inches(0.6), Inches(2.55), Inches(5.95), Inches(3.05), fill=WHITE, line=TEAL, line_w=Pt(1.5))
    rect(s, Inches(0.6), Inches(2.55), Inches(5.95), Inches(0.6), fill=TEAL)
    txt(s, Inches(0.85), Inches(2.55), Inches(5.5), Inches(0.6),
        [P("RAPIDO", 16, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(0.85), Inches(3.3), Inches(5.5), Inches(2.2),
        [bullet(p, 14) for p in rapido_pts], space_after=8)
    # Saarthi column
    rect(s, Inches(6.75), Inches(2.55), Inches(5.95), Inches(3.05), fill=WHITE, line=INDIGO, line_w=Pt(1.5))
    rect(s, Inches(6.75), Inches(2.55), Inches(5.95), Inches(0.6), fill=INDIGO)
    txt(s, Inches(7.0), Inches(2.55), Inches(5.5), Inches(0.6),
        [P("SAARTHI", 16, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(7.0), Inches(3.3), Inches(5.5), Inches(2.2),
        [bullet(p, 14) for p in saarthi_pts], space_after=8)
    # takeaway
    rect(s, Inches(0.6), Inches(5.85), Inches(12.1), Inches(0.85), fill=NAVY)
    txt(s, Inches(0.95), Inches(5.85), Inches(11.4), Inches(0.85),
        [[("Takeaway:  ", 15, TEAL, True), (takeaway, 15, WHITE, False)]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    footer(s, page)


reason_slide(
    1, "Reason 1 \u00b7 the decisive one", "It plays directly to end-to-end transactions",
    "Louie wins where there is a real transaction to finish by voice. Only one app has that.",
    ["Core journey IS a transaction: book \u2192 confirm \u2192 pay \u2192 ride",
     "Voice can carry the user from intent to a completed booking",
     "Showcases our headline differentiator in a single demo"],
    ["An awareness / education app \u2014 no purchase happens in-app",
     "Journeys end at \u201clearn more\u201d, not \u201ctransaction complete\u201d",
     "Little for an end-to-end voice flow to actually do"],
    "Rapido lets Louie demonstrate its #1 differentiator; Saarthi structurally cannot.", 9)

reason_slide(
    2, "Reason 2", "Riding is a low-deliberation command; investing is not",
    "People happily delegate simple, habitual actions to voice \u2014 but not high-stakes financial decisions.",
    ["\u201cBook an auto to home\u201d \u2014 instant, low-stakes, repeatable",
     "Few variables; the user already knows what they want",
     "Mistakes are cheap and easily corrected"],
    ["Investing needs research, comparison and deliberation",
     "Users want to read, verify and reflect before acting",
     "High stakes \u2014 people will not hand money decisions to voice"],
    "Voice automation adds most value for fast, habitual tasks \u2014 that is ride-booking, not investing.", 10)

reason_slide(
    3, "Reason 3", "The addressable market is far larger",
    "Reach multiplies our impact. Far more people ride than invest \u2014 and they do it far more often.",
    ["Rapido: ~35\u201340 lakh rides per day (tens of millions of monthly users)",
     "Ride-hailing is a daily, mass-market habit across cities & towns",
     "High frequency \u2014 multiple rides per user per week"],
    ["India active demat accounts \u2248 19 crore, but unique investors are far fewer",
     "Awareness-stage users invest rarely, if at all",
     "Even fewer invest impulsively \u2014 the low-frequency edge case"],
    "More users \u00d7 more frequency = far more voice interactions and a far bigger demo of our value.", 11)

reason_slide(
    4, "Reason 4", "Mobility is a universal need; voice makes it inclusive",
    "Voice should unlock essential daily life for the people who need it most.",
    ["Everyone needs to travel \u2014 a truly universal need",
     "Helps low-literacy, vernacular & visually-impaired users move independently",
     "Hands-busy / on-the-go moments make voice genuinely useful"],
    ["Investing is optional and niche, not a daily necessity",
     "Under-served groups rarely have surplus income to invest",
     "Accessibility upside is real but reaches far fewer people"],
    "Mobility delivers the broadest, most meaningful inclusion impact for a voice interface.", 12)

reason_slide(
    5, "Reason 5", "Vernacular voice widens Rapido's audience more",
    "Our multilingual edge matters most where the audience is large and not English-first.",
    ["Mass-market riders span many languages & literacy levels",
     "A vernacular voice layer removes the typing / English barrier",
     "Directly grows the usable user base for the app"],
    ["Awareness audience skews more urban, literate & English-comfortable",
     "Smaller base means smaller absolute language upside",
     "Language helps, but on a much smaller population"],
    "Our multilingual strength compounds with Rapido's scale \u2014 bigger base, bigger payoff.", 13)


# ----------------------------------------------------------------------------
# BY THE NUMBERS (quantitative evidence)
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Quantitative evidence", "The case, by the numbers", AMBER)
txt(s, Inches(0.6), Inches(1.5), Inches(12.1), Inches(0.45),
    [P("Scale and frequency overwhelmingly favour mobility over investing.", 15, SLATE)])

stats = [
    ("~35\u201340 lakh", "Rapido rides booked per day", "A transaction Louie can complete, millions of times daily", TEAL),
    ("Tens of\nmillions", "Monthly active Rapido users", "Huge base for vernacular, voice-led booking", TEAL),
    ("\u2248 80 crore+", "Smartphone users in India", "The reachable surface for a voice interface", INDIGO),
    ("\u2248 19 crore", "Active demat accounts (investing)", "And unique investors are far fewer still", RED),
    ("Small %\nof India", "Invests in markets at all", "Even fewer invest impulsively \u2014 voice's weak spot", RED),
    ("Daily", "Frequency of a commute vs. an investment", "Rides recur daily; investing is occasional", AMBER),
]
cw, ch, gx, gy = Inches(3.95), Inches(2.15), Inches(0.2), Inches(0.2)
x0, y0 = Inches(0.6), Inches(2.15)
for i, (big, label, sub, col) in enumerate(stats):
    x = x0 + (i % 3) * (cw + gx)
    y = y0 + (i // 3) * (ch + gy)
    rect(s, x, y, cw, ch, fill=LIGHT)
    rect(s, x, y, cw, Pt(6), fill=col)
    txt(s, x + Inches(0.25), y + Inches(0.2), cw - Inches(0.5), Inches(0.95),
        [P(big, 30, col, True)], line_spacing=0.9)
    txt(s, x + Inches(0.25), y + Inches(1.2), cw - Inches(0.5), Inches(0.4),
        [P(label, 13.5, NAVY, True)], line_spacing=1.0)
    txt(s, x + Inches(0.25), y + Inches(1.58), cw - Inches(0.5), Inches(0.5),
        [P(sub, 11.5, SLATE)], line_spacing=1.0)

txt(s, Inches(0.6), Inches(6.75), Inches(12.1), Inches(0.35),
    [[("Figures are approximate, directional public estimates \u2014 ", 10.5, SLATE, True),
      ("verify and cite exact, dated sources before the final presentation.", 10.5, SLATE, False)]])
footer(s)


# ----------------------------------------------------------------------------
# 14. SCORECARD
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Side by side", "Comparison scorecard (qualitative)", INDIGO)
rows = [
    ("Criterion", "Rapido", "Saarthi", True),
    ("End-to-end transaction fit", "High", "Low", False),
    ("Decision simplicity for voice", "High", "Low", False),
    ("Market reach / frequency", "High", "Low", False),
    ("Inclusion impact", "High", "Medium", False),
    ("Multilingual upside (absolute)", "High", "Low\u2013Med", False),
    ("Fit with Louie differentiators", "Strong", "Weak", False),
]
top = Inches(1.7)
rh = Inches(0.66)
c0, c1, c2 = Inches(0.6), Inches(7.0), Inches(9.85)
w0, w1, w2 = Inches(6.4), Inches(2.85), Inches(2.85)

def rating_color(v):
    v = v.lower()
    if v.startswith("high") or v.startswith("strong"):
        return TEAL
    if v.startswith("low") and "med" not in v or v.startswith("weak"):
        return RED
    return AMBER

for i, (crit, ra, sa, ishead) in enumerate(rows):
    y = top + i * rh
    if ishead:
        rect(s, c0, y, w0, rh, fill=NAVY)
        rect(s, c1, y, w1, rh, fill=TEAL)
        rect(s, c2, y, w2, rh, fill=INDIGO)
        txt(s, c0 + Inches(0.2), y, w0 - Inches(0.3), rh, [P(crit, 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        txt(s, c1, y, w1, rh, [P(ra, 15, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, c2, y, w2, rh, [P(sa, 15, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    else:
        bg = LIGHT if i % 2 else WHITE
        rect(s, c0, y, w0, rh, fill=bg)
        rect(s, c1, y, w1, rh, fill=bg)
        rect(s, c2, y, w2, rh, fill=bg)
        txt(s, c0 + Inches(0.2), y, w0 - Inches(0.3), rh, [P(crit, 14, DARKTEXT)], anchor=MSO_ANCHOR.MIDDLE)
        txt(s, c1, y, w1, rh, [P(ra, 14, rating_color(ra), True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, c2, y, w2, rh, [P(sa, 14, rating_color(sa), True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s)


# ----------------------------------------------------------------------------
# WEIGHTED SCORING MATRIX
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Decision analytics", "Weighted scoring matrix", TEAL)
txt(s, Inches(0.6), Inches(1.45), Inches(12.1), Inches(0.4),
    [P("Each criterion scored 0\u201310, weighted by importance to Louie's value. Totals are the weighted sums.", 14, SLATE)])

# (criterion, weight%, rapido, saarthi)
matrix = [
    ("Transaction fit (end-to-end)", 35, 10, 3),
    ("Market reach / frequency", 20, 9, 5),
    ("Inclusion impact", 15, 9, 6),
    ("Multilingual impact", 15, 9, 6),
    ("Ease of voice adoption", 15, 10, 4),
]
rapido_total = sum(w * r for _, w, r, _ in matrix) / 100.0
saarthi_total = sum(w * sa for _, w, _, sa in matrix) / 100.0

top = Inches(2.05)
rh = Inches(0.6)
cc = [Inches(0.6), Inches(6.3), Inches(8.0), Inches(9.85), Inches(11.7)]
cw_ = [Inches(5.6), Inches(1.6), Inches(1.75), Inches(1.75), Inches(1.0)]
heads = ["Criterion", "Weight", "Rapido", "Saarthi", ""]
# header row
for j, htext in enumerate(heads):
    fill = NAVY if j == 0 else (GREY if j == 1 else (TEAL if j == 2 else (INDIGO if j == 3 else NAVY)))
    rect(s, cc[j], top, cw_[j], rh, fill=fill)
    al = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
    xpad = Inches(0.2) if j == 0 else Inches(0)
    txt(s, cc[j] + xpad, top, cw_[j] - xpad, rh, [P(htext, 14, WHITE, True)], align=al, anchor=MSO_ANCHOR.MIDDLE)

for i, (crit, w, r, sa) in enumerate(matrix):
    y = top + (i + 1) * rh
    bg = LIGHT if i % 2 else WHITE
    for j in range(5):
        rect(s, cc[j], y, cw_[j], rh, fill=bg)
    txt(s, cc[0] + Inches(0.2), y, cw_[0] - Inches(0.3), rh, [P(crit, 13.5, DARKTEXT)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, cc[1], y, cw_[1], rh, [P(f"{w}%", 13.5, SLATE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, cc[2], y, cw_[2], rh, [P(str(r), 14, TEAL, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, cc[3], y, cw_[3], rh, [P(str(sa), 14, INDIGO, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    win = "\u2713" if r >= sa else ""
    txt(s, cc[4], y, cw_[4], rh, [P(win, 14, TEAL, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# totals row
y = top + (len(matrix) + 1) * rh
rect(s, cc[0], y, cw_[0] + cw_[1], rh, fill=NAVY)
txt(s, cc[0] + Inches(0.2), y, cw_[0] + cw_[1] - Inches(0.3), rh,
    [P("Weighted total (out of 10)", 14, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
rect(s, cc[2], y, cw_[2], rh, fill=TEAL)
txt(s, cc[2], y, cw_[2], rh, [P(f"{rapido_total:.1f}", 16, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, cc[3], y, cw_[3], rh, fill=INDIGO)
txt(s, cc[3], y, cw_[3], rh, [P(f"{saarthi_total:.1f}", 16, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, cc[4], y, cw_[4], rh, fill=TEAL)
txt(s, cc[4], y, cw_[4], rh, [P("\u2713", 14, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

txt(s, Inches(0.6), y + Inches(0.8), Inches(12), Inches(0.4),
    [[("Verdict:  ", 15, NAVY, True),
      (f"Rapido {rapido_total:.1f} vs Saarthi {saarthi_total:.1f} \u2014 a decisive margin, driven mainly by transaction fit.",
       15, DARKTEXT, False)]])
footer(s)


# ----------------------------------------------------------------------------
# COMMERCIAL / ROI ANGLE
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Commercial lens", "Why Rapido is the better bet for Louie's business", TEAL)
txt(s, Inches(0.6), Inches(1.5), Inches(12.1), Inches(0.45),
    [P("If a stakeholder asks \u201cHow does Louie benefit commercially?\u201d \u2014 the answer is stronger with Rapido.", 15, SLATE)])

roi = [
    ("High daily usage", "Recurring rides mean continuous, sticky engagement with the voice layer.", TEAL),
    ("Frequent voice interactions", "More sessions = more data to tune models and prove reliability.", INDIGO),
    ("Strong demo value", "\u201cBook a ride end-to-end by voice\u201d is an instantly compelling sales demo.", AMBER),
    ("Easier case-study generation", "High volume yields measurable adoption, retention and CSAT proof points.", TEAL),
    ("Faster proof-of-concept", "A simple, bounded flow ships quickly and de-risks the engagement.", INDIGO),
    ("Expansion runway", "Same playbook extends to food delivery, logistics & broader mobility.", AMBER),
]
cw, ch, gx, gy = Inches(3.95), Inches(2.0), Inches(0.2), Inches(0.2)
x0, y0 = Inches(0.6), Inches(2.15)
for i, (h, b, col) in enumerate(roi):
    x = x0 + (i % 3) * (cw + gx)
    y = y0 + (i // 3) * (ch + gy)
    rect(s, x, y, cw, ch, fill=LIGHT)
    rect(s, x, y, Inches(0.12), ch, fill=col)
    txt(s, x + Inches(0.3), y + Inches(0.22), cw - Inches(0.5), Inches(0.7),
        [P(h, 15.5, col, True)], line_spacing=1.0)
    txt(s, x + Inches(0.3), y + Inches(0.95), cw - Inches(0.5), Inches(0.95),
        [P(b, 12.5, DARKTEXT)], line_spacing=1.03)

txt(s, Inches(0.6), Inches(6.75), Inches(12.1), Inches(0.35),
    [[("Bottom line: ", 12, NAVY, True),
      ("Rapido offers faster time-to-value, a better showcase, and a clear path to adjacent transactional categories.", 12, SLATE, False)]])
footer(s)


# ----------------------------------------------------------------------------
# 15. RISKS & REBUTTALS
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Stress-testing the call", "Counter-points & our response", AMBER)
qa = [
    ("\u201cSaarthi has SEBI / government backing \u2014 great for credibility.\u201d",
     "Valuable, but it does not create an end-to-end transaction. We can pursue it later once capacity frees up, without spending our scarce slot on a weak differentiator fit."),
    ("\u201cFin-services is a high-value, premium segment.\u201d",
     "True for trading apps that move money \u2014 not for an awareness app. The transactional fintech opportunity is a separate, future conversation."),
    ("\u201cRapido is a crowded mobility space.\u201d",
     "The app is crowded; voice-completed booking is not. Our differentiator is the experience, not the category."),
    ("\u201cAccessibility matters for investing too.\u201d",
     "Agreed \u2014 but mobility reaches a far larger, more universal population, so the same effort delivers more inclusion impact."),
]
y = Inches(1.7)
for q, a in qa:
    rect(s, Inches(0.6), y, Inches(12.1), Inches(1.18), fill=LIGHT)
    rect(s, Inches(0.6), y, Pt(5), Inches(1.18), fill=AMBER)
    txt(s, Inches(0.9), y + Inches(0.12), Inches(11.5), Inches(0.45),
        [[("Concern:  ", 13.5, AMBER, True), (q, 13.5, DARKTEXT, True)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(0.9), y + Inches(0.56), Inches(11.5), Inches(0.55),
        [[("Response:  ", 13.5, TEAL, True), (a, 13.5, DARKTEXT, False)]], line_spacing=1.0)
    y += Inches(1.28)
footer(s)


# ----------------------------------------------------------------------------
# RISKS & MITIGATION (delivery risks)
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Practical delivery", "Key risks & mitigations", RED)
txt(s, Inches(0.6), Inches(1.5), Inches(12.1), Inches(0.4),
    [P("Voice-enabling a live transaction carries real-world risks \u2014 each has a practical mitigation.", 15, SLATE)])

# table header
hy = Inches(2.1)
rh = Inches(0.62)
xc1, xc2 = Inches(0.6), Inches(6.55)
wc1, wc2 = Inches(5.75), Inches(6.15)
rect(s, xc1, hy, wc1, rh, fill=RED)
rect(s, xc2, hy, wc2, rh, fill=TEAL)
txt(s, xc1 + Inches(0.2), hy, wc1 - Inches(0.3), rh, [P("Risk", 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
txt(s, xc2 + Inches(0.2), hy, wc2 - Inches(0.3), rh, [P("Mitigation", 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)

risks = [
    ("Incorrect pickup / drop location captured via voice",
     "Explicit read-back & confirmation step before booking; map pin preview."),
    ("Accent, dialect & language variation",
     "Multilingual model tuning; per-language acoustic & vocabulary adaptation."),
    ("Noisy outdoor / on-the-move environment",
     "Hybrid voice + touch fallback; noise-robust ASR and re-prompting."),
    ("Mis-triggered payment or wrong fare",
     "Confirm amount aloud; require explicit \u201cyes\u201d before charging."),
    ("Recognition failure / user frustration",
     "Graceful fallback to standard UI; never block the core booking flow."),
]
ry = hy + rh
rrh = Inches(0.7)
for i, (risk, mit) in enumerate(risks):
    bg = LIGHT if i % 2 else WHITE
    rect(s, xc1, ry, wc1, rrh, fill=bg)
    rect(s, xc2, ry, wc2, rrh, fill=bg)
    txt(s, xc1 + Inches(0.2), ry, wc1 - Inches(0.35), rrh, [P(risk, 13, DARKTEXT)], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    txt(s, xc2 + Inches(0.2), ry, wc2 - Inches(0.35), rrh, [P(mit, 13, DARKTEXT)], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    ry += rrh

txt(s, Inches(0.6), ry + Inches(0.12), Inches(12.1), Inches(0.4),
    [[("Note: ", 12, NAVY, True),
      ("confirmation-first design keeps every voice transaction safe, reversible and trustworthy.", 12, SLATE, False)]])
footer(s)


# ----------------------------------------------------------------------------
# 16. RECOMMENDATION & NEXT STEPS
# ----------------------------------------------------------------------------
s = add_slide()
rect(s, 0, 0, EMU_W, EMU_H, fill=NAVY)
rect(s, 0, 0, Inches(0.35), EMU_H, fill=TEAL)
txt(s, Inches(0.9), Inches(0.7), Inches(11), Inches(0.5),
    [P("RECOMMENDATION", 14, TEAL, True)])
txt(s, Inches(0.9), Inches(1.25), Inches(11.5), Inches(1.1),
    [[("Onboard ", 34, WHITE, True), ("Rapido", 34, TEAL, True),
      (" \u2014 it is where Louie wins.", 34, WHITE, True)]])
txt(s, Inches(0.9), Inches(2.45), Inches(11.5), Inches(0.9),
    [P("It is the only option whose core journey is an end-to-end, multilingual transaction \u2014 our single strongest differentiator \u2014 at the largest scale and with the broadest inclusion impact.",
       16, RGBColor(0xCF,0xD3,0xE6))], line_spacing=1.1)

txt(s, Inches(0.9), Inches(3.7), Inches(11), Inches(0.4),
    [P("Suggested next steps", 17, WHITE, True)])
steps = [
    "Confirm scope with Rapido: voice-enable the end-to-end booking flow first",
    "Prioritise top vernacular languages by rider base for the voice layer",
    "Define a flagship demo: \u201cBook a ride, by voice, in your language\u201d",
    "Park Saarthi as a warm lead; revisit when capacity frees up or a transactional fintech use case emerges",
]
y = Inches(4.2)
for i, st in enumerate(steps, 1):
    c = rect(s, Inches(0.95), y + Inches(0.05), Inches(0.5), Inches(0.5), fill=TEAL, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.95), y + Inches(0.05), Inches(0.5), Inches(0.5),
        [P(str(i), 16, NAVY, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.65), y, Inches(11), Inches(0.6),
        [P(st, 15, RGBColor(0xE5,0xE7,0xF2))], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.62)
footer(s, 16)


# ----------------------------------------------------------------------------
# 17. APPENDIX
# ----------------------------------------------------------------------------
s = add_slide()
header(s, "Appendix", "References & download links", SLATE)
txt(s, Inches(0.6), Inches(1.7), Inches(12), Inches(0.4),
    [P("Apps evaluated", 16, NAVY, True)])
links = [
    ("Saarthi (Android)", "play.google.com/store/apps/details?id=com.sebi.invapp"),
    ("Rapido (Android)", "play.google.com/store/apps/details?id=com.rapido.passenger"),
    ("Saarthi (iOS)", "apps.apple.com/in/app/saa\u20b9thi/id1589426387"),
    ("Rapido (iOS)", "apps.apple.com/in/app/rapido-bike-taxi-auto-cabs/id1198464606"),
]
y = Inches(2.2)
for name, url in links:
    txt(s, Inches(0.8), y, Inches(3.0), Inches(0.35), [P(name, 13, DARKTEXT, True)])
    txt(s, Inches(3.9), y, Inches(8.7), Inches(0.35), [P(url, 13, INDIGO)])
    y += Inches(0.45)

txt(s, Inches(0.6), Inches(4.3), Inches(12), Inches(0.4),
    [P("Louie reference demos (for context)", 16, NAVY, True)])
demos = [
    ("Rapido booking with Louie", "youtube.com/watch?v=-W22qTvvFS4"),
    ("UPI transfer (showcased at GFF)", "youtube.com/shorts/fSic13x9Do4"),
    ("Uber booking: Louie vs Google Assistant vs Siri", "see comparison links in brief"),
]
y = Inches(4.8)
for name, url in demos:
    txt(s, Inches(0.8), y, Inches(4.5), Inches(0.35), [P(name, 13, DARKTEXT, True)])
    txt(s, Inches(5.4), y, Inches(7.2), Inches(0.35), [P(url, 13, INDIGO)])
    y += Inches(0.45)

txt(s, Inches(0.6), Inches(6.45), Inches(12), Inches(0.5),
    [[("To finalise: ", 12, SLATE, True),
      ("add real screenshots on slides 6 & 7, fill in author/date on the title slide, and tailor any India-market stats you wish to cite.", 12, SLATE, False)]])
footer(s, 17)


# ----------------------------------------------------------------------------
out = "Louie_Voice_SDK_Rapido_vs_Saarthi.pptx"
prs.save(out)
print(f"Saved {out} with {len(prs.slides._sldIdLst)} slides.")
