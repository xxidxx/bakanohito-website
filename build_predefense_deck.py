# -*- coding: utf-8 -*-
"""Pre-defense thesis deck template — NEU standard. Pertamax Green 95."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ---- constants ----
W, H = Inches(20), Inches(11.25)
ML = Inches(1.15)                 # left margin
CW = Inches(20 - 2*1.15)          # content width 17.71
INK      = RGBColor(0x16,0x19,0x1C)
GREEN    = RGBColor(0x1A,0x6E,0x4F)
GRAY     = RGBColor(0x4A,0x4F,0x54)
DGRAY    = RGBColor(0x34,0x38,0x3C)
CAP      = RGBColor(0x5B,0x62,0x5F)
MUTED    = RGBColor(0x8A,0x90,0x99)
RED      = RGBColor(0xC0,0x49,0x2F)
HAIR     = RGBColor(0xE3,0xE6,0xE9)
BOXLINE  = RGBColor(0xD8,0xDC,0xE0)
GHOST    = RGBColor(0xEE,0xF0,0xF2)
LGREEN   = RGBColor(0xBF,0xE0,0xD2)
WHITE    = RGBColor(0xFF,0xFF,0xFF)
LOGO = r"C:\Users\idari\Downloads\Hermes\neu_logo_cropped.png"  # cropped asset, content ratio 2.712
FONT = "Arial"

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def rect(s, l, t, w, h, fill=None, line=None, line_w=0.43):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.shadow.inherit = False
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(line_w)
    return sh

def tb(s, l, t, w, h, runs, size=18, color=GRAY, bold=False, align=PP_ALIGN.LEFT,
       anchor=MSO_ANCHOR.TOP, space_after=6, line_spacing=1.0):
    """runs: str, or list of paragraphs; each paragraph str or list of (text, dict) runs."""
    box = s.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str): runs = [runs]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing: p.line_spacing = line_spacing
        if isinstance(para, str): para = [(para, {})]
        for text, fmt in para:
            r = p.add_run(); r.text = text
            r.font.name = FONT
            r.font.size = Pt(fmt.get("size", size))
            r.font.bold = fmt.get("bold", bold)
            r.font.color.rgb = fmt.get("color", color)
    return box

def logo_topright(s):
    s.shapes.add_picture(LOGO, Inches(16.69), Inches(0.90), Inches(0.80*2.712), Inches(0.80))

def hairline(s, t=Inches(2.16), l=ML, w=CW):
    rect(s, l, t, w, Pt(1.2), fill=HAIR)

def header(s, kicker, headline):
    tb(s, ML, Inches(0.88), Inches(14), Inches(0.4), kicker.upper(), size=18, color=GREEN, bold=True)
    tb(s, ML, Inches(1.31), Inches(14), Inches(0.8), headline, size=39, color=INK, bold=True)
    hairline(s)
    logo_topright(s)

def green_tick(s, l, t):
    rect(s, l, t, Inches(0.46), Inches(0.03), fill=GREEN)

def stat_card(s, l, t, w, h, number, caption, num_color=INK, num_size=45):
    rect(s, l, t, w, h, fill=None, line=BOXLINE)
    tb(s, l+Inches(0.3), t+Inches(0.25), w-Inches(0.6), Inches(1.0), number, size=num_size, color=num_color, bold=True)
    tb(s, l+Inches(0.3), t+h-Inches(1.15), w-Inches(0.6), Inches(1.0), caption, size=18, color=CAP, line_spacing=1.05)

def numbered(s, l, t, w, num, title, desc, sep=True):
    tb(s, l, t, Inches(0.9), Inches(0.6), num, size=25.5, color=GREEN, bold=True)
    tb(s, l+Inches(1.0), t, w-Inches(1.0), Inches(0.55), title, size=24, color=INK, bold=True)
    d = tb(s, l+Inches(1.0), t+Inches(0.5), w-Inches(1.0), Inches(1.0), desc, size=18.75, color=GRAY, line_spacing=1.1)
    if sep:
        rect(s, l, t+Inches(1.7), w, Pt(1.0), fill=HAIR)

def feature_block(s, l, t, w, h, title, body):
    rect(s, l, t, w, h, fill=GREEN)
    tb(s, l+Inches(0.4), t+Inches(0.3), w-Inches(0.8), Inches(0.6), title.upper(), size=18, color=LGREEN, bold=True)
    tb(s, l+Inches(0.4), t+Inches(0.85), w-Inches(0.8), h-Inches(1.1), body, size=21, color=WHITE, line_spacing=1.1)

def side_callout(s, l, t, w, label, body, bar=GREEN):
    rect(s, l, t, Inches(0.04), Inches(1.5), fill=bar)
    tb(s, l+Inches(0.25), t, w-Inches(0.25), Inches(0.4), label.upper(), size=18, color=MUTED, bold=True)
    tb(s, l+Inches(0.25), t+Inches(0.45), w-Inches(0.25), Inches(1.4), body, size=19, color=GRAY, line_spacing=1.1)

# =========================================================
# 1. TITLE
# =========================================================
s = slide()
s.shapes.add_picture(LOGO, ML, Inches(0.55), Inches(0.95*2.712), Inches(0.95))
tb(s, Inches(12.5), Inches(0.9), Inches(6.3), Inches(0.4),
   "MBA THESIS PRE-DEFENSE · [MONTH YEAR]", size=18, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)
hairline(s, t=Inches(2.3))
green_tick(s, ML, Inches(3.3))
tb(s, ML, Inches(3.55), CW, Inches(0.5), "GRADUATE SCHOOL OF BUSINESS · MBA PROGRAM",
   size=18, color=GREEN, bold=True)
tb(s, ML, Inches(4.35), CW, Inches(2.8), [
    [("Marketing Strategy of ", {}), ("Pertamax Green 95", {"color": GREEN}), (":", {})],
    [("Pre-Defense Proposal", {})],
], size=52, color=INK, bold=True, line_spacing=1.05)
tb(s, ML, Inches(6.6), CW, Inches(0.5),
   "Research on the Marketing Strategy of Pertamina's Pertamax Green 95 in the Indonesian Market",
   size=21, color=CAP)
hairline(s, t=Inches(8.3))
tb(s, ML, Inches(8.7), Inches(8), Inches(1.6), [
    [("Candidate:  ", {"color": MUTED, "bold": True}), ("Id Arifandi  ·  [NIM]", {"color": DGRAY, "bold": True})],
    [("Supervisor:  ", {"color": MUTED, "bold": True}), ("[Prof. Dr. Name]", {"color": DGRAY})],
    [("Date:  ", {"color": MUTED, "bold": True}), ("[DD Month 2026]", {"color": DGRAY})],
], size=18)

# =========================================================
# helper: SECTION DIVIDER
# =========================================================
def divider(num, part, title, subtitle):
    s = slide()
    tb(s, Inches(11.5), Inches(2.2), Inches(8), Inches(4), num, size=150, color=GHOST, bold=True, align=PP_ALIGN.RIGHT)
    green_tick(s, ML, Inches(4.15))
    tb(s, ML, Inches(4.4), CW, Inches(0.5), part.upper(), size=18, color=GREEN, bold=True)
    tb(s, ML, Inches(5.0), CW, Inches(1.4), title, size=63, color=INK, bold=True)
    tb(s, ML, Inches(6.5), CW, Inches(0.6), subtitle, size=22.5, color=CAP)
    return s

# =========================================================
# 2. OUTLINE
# =========================================================
s = slide()
header(s, "Pre-Defense Proposal", "Presentation outline")
items = [
    ("01", "Research Background", "Industry context, green fuel transition, and why Pertamax Green 95 matters."),
    ("02", "Problem Statement & Objectives", "Research questions, objectives, scope, and significance."),
    ("03", "Literature Review", "Theoretical foundation: STP, marketing mix (7P), consumer behavior."),
    ("04", "Research Methodology", "Design, data sources, sampling, and analysis framework."),
    ("05", "Preliminary Findings & Plan", "Expected outcomes, contribution, and research timeline."),
]
y = 2.6
for num, t, d in items:
    numbered(s, ML, Inches(y), Inches(10.8), num, t, d)
    y += 1.85
feature_block(s, Inches(12.7), Inches(2.6), Inches(6.16), Inches(8.0),
    "Goal of this defense",
    "Obtain committee feedback and approval to proceed with the full research on Pertamax Green 95's marketing strategy for the Indonesian market.")

# =========================================================
# 3. SECTION DIVIDER — PART ONE
# =========================================================
divider("01", "Part One", "Research Background", "Context of Indonesia's fuel market and green energy transition")

# =========================================================
# 4. BACKGROUND — STAT GRID
# =========================================================
s = slide()
header(s, "Part One · Research Background", "Fuel market shifting toward cleaner energy")
cw3 = Inches(5.70); gap = Inches(0.30)
stats = [
    ("[XX]%", "Fuel demand growth in Indonesia, [year] — transport sector dominates consumption.", INK),
    ("[XX]%", "Pertalite market share — subsidized fuel still dominates retail volume.", RED),
    ("RP [X]T", "Government target / investment for green energy transition by [year].", GREEN),
]
for i, (n, c, col) in enumerate(stats):
    stat_card(s, ML + i*(cw3+gap), Inches(2.6), cw3, Inches(2.8), n, c, num_color=col)
tb(s, ML, Inches(5.9), CW, Inches(0.5), "WHY THIS MATTERS", size=18, color=MUTED, bold=True)
tb(s, ML, Inches(6.4), CW, Inches(2.5), [
    [("Pertamina launched ", {}), ("Pertamax Green 95", {"bold": True, "color": INK}),
     (" as a bioethanol-blended gasoline (E5) to bridge the transition toward low-emission fuels in the Indonesian market.", {})],
    [("Adoption remains below target — the core driver of this research.", {"bold": True, "color": GREEN})],
], size=21, line_spacing=1.15, space_after=12)

# =========================================================
# 5. PROBLEM STATEMENT
# =========================================================
s = slide()
header(s, "Part One · Problem Statement", "Adoption lags behind launch targets")
side_callout(s, ML, Inches(2.7), Inches(8.4), "Core problem",
    "Despite national launch and Pertamina's brand strength, Pertamax Green 95 sales volume reaches only [X]% of target in [region/period].", bar=RED)
side_callout(s, ML, Inches(4.9), Inches(8.4), "Knowledge gap",
    "No prior study has examined the marketing strategy of bioethanol-blended gasoline in Indonesia using an integrated STP + 7P framework.", bar=GREEN)
y = 2.7
tb(s, Inches(10.4), Inches(2.7), Inches(8.4), Inches(0.4), "RESEARCH QUESTIONS", size=18, color=MUTED, bold=True)
rqs = [
    "RQ1 — How is the current marketing mix (7P) of Pertamax Green 95 implemented?",
    "RQ2 — What internal and external factors shape its market performance?",
    "RQ3 — What marketing strategy should Pertamina adopt to increase adoption?",
]
for i, q in enumerate(rqs):
    tb(s, Inches(10.4), Inches(3.25)+i*Inches(1.05), Inches(8.4), Inches(1.0), q, size=19, color=GRAY, line_spacing=1.08)
    if i < 2: rect(s, Inches(10.4), Inches(4.12)+i*Inches(1.05), Inches(8.4), Pt(1.0), fill=HAIR)
tb(s, ML, Inches(7.1), CW, Inches(1.8), [
    [("Objectives:  ", {"bold": True, "color": INK}),
     ("(1) evaluate the current marketing mix; (2) diagnose market barriers via SWOT/PESTLE; ", {}),
     ("(3) formulate a data-driven marketing strategy", {"bold": True, "color": GREEN}), (" for Pertamax Green 95.", {})],
], size=21, line_spacing=1.15)

# =========================================================
# 6. SECTION DIVIDER — PART TWO
# =========================================================
divider("02", "Part Two", "Literature Review", "Theoretical foundation and conceptual framework")

# =========================================================
# 7. LITERATURE REVIEW
# =========================================================
s = slide()
header(s, "Part Two · Literature Review", "Three pillars support the analysis")
cw3 = Inches(5.70); gap = Inches(0.30)
cards = [
    ("STP", "Segmentation · Targeting · Positioning", "Identifying priority consumer segments for green gasoline and sharpening Pertamax Green 95's positioning vs Pertalite/Pertamax.  (Kotler & Keller, [year])"),
    ("7P Marketing Mix", "Product · Price · Place · Promotion · People · Process · Physical Evidence", "Diagnosing the current service-dominant marketing mix of a fuel retail product.  (Booms & Bitner, [year])"),
    ("Consumer Behavior", "Awareness · Perception · Adoption", "Explaining green-fuel adoption barriers: price sensitivity, perceived quality, environmental concern.  ([Author, year])"),
]
for i, (t, sub, d) in enumerate(cards):
    l = ML + i*(cw3+gap)
    rect(s, l, Inches(2.6), cw3, Inches(6.6), fill=None, line=BOXLINE)
    rect(s, l, Inches(2.6), cw3, Inches(0.06), fill=GREEN)
    tb(s, l+Inches(0.3), Inches(2.95), cw3-Inches(0.6), Inches(0.6), t, size=25.5, color=INK, bold=True)
    tb(s, l+Inches(0.3), Inches(3.6), cw3-Inches(0.6), Inches(0.9), sub.upper(), size=14, color=GREEN, bold=True, line_spacing=1.1)
    tb(s, l+Inches(0.3), Inches(4.7), cw3-Inches(0.6), Inches(4.2), d, size=18.75, color=GRAY, line_spacing=1.12)
tb(s, ML, Inches(9.6), CW, Inches(0.9), [
    [("Conceptual framework: ", {}),
     ("Internal & external analysis (SWOT, PESTLE, Porter)  →  STP  →  7P strategy formulation", {"bold": True, "color": INK})],
], size=21)

# =========================================================
# 8. SECTION DIVIDER — PART THREE
# =========================================================
divider("03", "Part Three", "Research Methodology", "Design, data, and analysis plan")

# =========================================================
# 9. METHODOLOGY
# =========================================================
s = slide()
header(s, "Part Three · Methodology", "Mixed methods: internal data plus survey")
rows = [
    ("Research design", "Descriptive–explanatory, mixed methods (qualitative + quantitative)"),
    ("Unit of analysis", "Pertamax Green 95 marketing program, Pertamina [Region/Area]"),
    ("Primary data", "Consumer survey (n = [X], purposive sampling) + in-depth interviews with [X] Pertamina marketing staff"),
    ("Secondary data", "Internal sales volume, pricing and distribution data; industry reports; ESDM regulations"),
    ("Analysis tools", "SWOT + QSPM (strategy prioritization), PESTLE, Porter's Five Forces, descriptive statistics"),
    ("Research period", "[Month Year] – [Month Year]"),
]
y = 2.6
for i, (k, v) in enumerate(rows):
    tb(s, ML, Inches(y), Inches(4.4), Inches(0.5), k.upper(), size=16, color=MUTED, bold=True)
    tb(s, Inches(5.8), Inches(y-0.08), Inches(13.0), Inches(1.0),
       [[(v, {})]], size=20, color=DGRAY, line_spacing=1.05)
    rect(s, ML, Inches(y+0.72), CW, Pt(1.0), fill=HAIR)
    y += 1.22

# =========================================================
# 10. PRELIMINARY FINDINGS (placeholder)
# =========================================================
s = slide()
header(s, "Part Three · Preliminary Observations", "Early signals from internal and field data")
cw2 = Inches(8.70); gap = Inches(0.31)
stat_card(s, ML, Inches(2.6), cw2, Inches(3.0), "[X.X] M", "Pertamax Green 95 sales volume (KL) in [period] — versus target of [X.X] M KL.", num_color=RED)
stat_card(s, ML+cw2+gap, Inches(2.6), cw2, Inches(3.0), "[XX]%", "Consumer awareness of Pertamax Green 95 in preliminary field check (n=[X]).", num_color=INK)
tb(s, ML, Inches(6.1), CW, Inches(0.4), "WORKING HYPOTHESES", size=18, color=MUTED, bold=True)
tb(s, ML, Inches(6.6), CW, Inches(2.4), [
    [("H1: ", {"bold": True, "color": GREEN}), ("Price premium over Pertalite is the primary adoption barrier for price-sensitive segments.", {})],
    [("H2: ", {"bold": True, "color": GREEN}), ("Low promotion intensity and limited distribution coverage suppress trial and repeat purchase.", {})],
], size=21, line_spacing=1.15, space_after=10)

# =========================================================
# 11. TIMELINE + CONTRIBUTION
# =========================================================
s = slide()
header(s, "Part Three · Plan", "Timeline and expected contribution")
tb(s, ML, Inches(2.7), Inches(5), Inches(0.4), "TIMELINE", size=18, color=MUTED, bold=True)
phases = [
    ("Data collection & survey", "[Mon–Mon 2026]"),
    ("Analysis (SWOT, QSPM, STP)", "[Mon–Mon 2026]"),
    ("Strategy formulation & writing", "[Mon–Mon 2026]"),
    ("Thesis defense", "[Month 2026]"),
]
y = 3.3
for i, (p, d) in enumerate(phases):
    green_tick(s, ML, Inches(y+0.12))
    tb(s, ML+Inches(0.75), Inches(y), Inches(7.5), Inches(0.5), p, size=21, color=DGRAY, bold=True)
    tb(s, ML+Inches(0.75), Inches(y+0.45), Inches(7.5), Inches(0.4), d, size=17, color=CAP)
    y += 1.35
feature_block(s, Inches(9.9), Inches(2.7), Inches(8.95), Inches(6.3),
    "Expected contribution",
    "Academic — first integrated STP + 7P study of bioethanol-blended gasoline in Indonesia.\n\nPractical — an actionable, prioritized marketing strategy roadmap for Pertamina to accelerate Pertamax Green 95 adoption and support the national energy transition.")

# =========================================================
# 12. CLOSING
# =========================================================
s = slide()
hairline(s, t=Inches(2.3))
green_tick(s, ML, Inches(4.0))
tb(s, ML, Inches(4.35), CW, Inches(1.6), "Q & A", size=96, color=INK, bold=True)
tb(s, ML, Inches(6.2), CW, Inches(0.7), "Thank you — feedback and suggestions are welcome.",
   size=28, color=GRAY)
tb(s, ML, Inches(7.6), CW, Inches(0.5),
   "Research on the Marketing Strategy of Pertamina's Pertamax Green 95 in the Indonesian Market",
   size=17, color=MUTED)
hairline(s, t=Inches(9.3))
s.shapes.add_picture(LOGO, ML, Inches(9.6), Inches(0.70*2.712), Inches(0.70))
tb(s, Inches(12.5), Inches(9.85), Inches(6.3), Inches(0.4), "Id Arifandi · MBA Program",
   size=16, color=MUTED, align=PP_ALIGN.RIGHT)

out = r"C:\Users\idari\Downloads\Hermes\PertamaxGreen95_PreDefense_Template.pptx"
prs.save(out)
print("saved", out, "slides:", len(prs.slides.__iter__.__self__._sldIdLst) if False else len(prs.slides._sldIdLst))
