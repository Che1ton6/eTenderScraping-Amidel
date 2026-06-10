#!/usr/bin/env python3
"""Generates the Amidel eTender Scraper product presentation."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Brand colours ──────────────────────────────────────────────────────────────
NAVY        = RGBColor(0x1C, 0x38, 0x80)
NAVY_LIGHT  = RGBColor(0x25, 0x44, 0x99)
ORANGE      = RGBColor(0xF5, 0xA0, 0x00)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE   = RGBColor(0xF4, 0xF6, 0xFA)
DARK_GRAY   = RGBColor(0x33, 0x33, 0x44)
MID_GRAY    = RGBColor(0x66, 0x66, 0x77)
LIGHT_GRAY  = RGBColor(0xE0, 0xE4, 0xEE)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def new_prs():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


# ── Low-level helpers ──────────────────────────────────────────────────────────

def add_rect(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_width_pt=0):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.line.fill.background()
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if line_rgb and line_width_pt:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=DARK_GRAY,
                 align=PP_ALIGN.LEFT, font_name="Segoe UI",
                 italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def add_para(tf, text, font_size=14, bold=False, color=DARK_GRAY,
             align=PP_ALIGN.LEFT, font_name="Segoe UI", italic=False,
             space_before_pt=0):
    from pptx.util import Pt as _Pt
    p = tf.add_paragraph()
    p.alignment = align
    if space_before_pt:
        p.space_before = _Pt(space_before_pt)
    run = p.add_run()
    run.text = text
    run.font.size = _Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return p


def navy_header_band(slide, title, subtitle=None):
    """Full-width navy header band with title + optional subtitle."""
    band_h = Inches(1.45)
    add_rect(slide, 0, 0, SLIDE_W, band_h, fill_rgb=NAVY)
    # Orange accent strip
    add_rect(slide, 0, band_h - Inches(0.06), SLIDE_W, Inches(0.06), fill_rgb=ORANGE)

    add_text_box(slide, title,
                 left=Inches(0.55), top=Inches(0.18),
                 width=Inches(12.2), height=Inches(0.7),
                 font_size=28, bold=True, color=WHITE,
                 align=PP_ALIGN.LEFT)
    if subtitle:
        add_text_box(slide, subtitle,
                     left=Inches(0.55), top=Inches(0.82),
                     width=Inches(12.2), height=Inches(0.4),
                     font_size=14, color=ORANGE, italic=True,
                     align=PP_ALIGN.LEFT)


def content_area_top():
    return Inches(1.65)


def slide_footer(slide, page_num, total):
    add_rect(slide, 0, SLIDE_H - Inches(0.32), SLIDE_W, Inches(0.32), fill_rgb=NAVY)
    add_text_box(slide, "Amidel (Pty) Ltd  |  Beyond Technology  |  eTender Scraper",
                 left=Inches(0.4), top=SLIDE_H - Inches(0.31),
                 width=Inches(10), height=Inches(0.3),
                 font_size=9, color=OFF_WHITE, align=PP_ALIGN.LEFT)
    add_text_box(slide, f"{page_num} / {total}",
                 left=Inches(12.5), top=SLIDE_H - Inches(0.31),
                 width=Inches(0.7), height=Inches(0.3),
                 font_size=9, color=ORANGE, align=PP_ALIGN.RIGHT)


def bullet_card(slide, left, top, width, height, heading, bullets,
                heading_size=15, bullet_size=12.5, icon=None):
    """A card with an optional icon, heading, and bullet list."""
    add_rect(slide, left, top, width, height,
             fill_rgb=OFF_WHITE, line_rgb=LIGHT_GRAY, line_width_pt=0.75)
    # Orange left accent
    add_rect(slide, left, top, Inches(0.055), height, fill_rgb=ORANGE)

    tx_left = left + Inches(0.18)
    tx_width = width - Inches(0.22)

    txBox = slide.shapes.add_textbox(tx_left, top + Inches(0.14), tx_width, height - Inches(0.14))
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = (icon + "  " if icon else "") + heading
    run.font.size = Pt(heading_size)
    run.font.bold = True
    run.font.color.rgb = NAVY
    run.font.name = "Segoe UI"

    for b in bullets:
        add_para(tf, "• " + b, font_size=bullet_size, color=DARK_GRAY)


# ── SLIDE BUILDERS ─────────────────────────────────────────────────────────────

TOTAL_SLIDES = 10


def slide_01_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Full navy background
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=NAVY)
    # Orange stripe at bottom
    add_rect(slide, 0, SLIDE_H - Inches(0.5), SLIDE_W, Inches(0.5), fill_rgb=ORANGE)
    # Decorative right panel
    add_rect(slide, SLIDE_W - Inches(3.4), 0, Inches(3.4), SLIDE_H - Inches(0.5),
             fill_rgb=NAVY_LIGHT)
    # Orange accent line separating panels
    add_rect(slide, SLIDE_W - Inches(3.4), 0, Inches(0.055), SLIDE_H - Inches(0.5),
             fill_rgb=ORANGE)

    # Product label chip
    add_rect(slide, Inches(0.7), Inches(1.7), Inches(2.8), Inches(0.4), fill_rgb=ORANGE)
    add_text_box(slide, "INTERNAL TOOL  ·  v2.0",
                 left=Inches(0.72), top=Inches(1.72),
                 width=Inches(2.8), height=Inches(0.38),
                 font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Main title
    add_text_box(slide, "Amidel\neTender Scraper",
                 left=Inches(0.7), top=Inches(2.2),
                 width=Inches(9.5), height=Inches(2.0),
                 font_size=52, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Tagline
    add_text_box(slide, "Automate. Aggregate. Act.",
                 left=Inches(0.7), top=Inches(4.25),
                 width=Inches(9), height=Inches(0.5),
                 font_size=20, bold=False, color=ORANGE, italic=True,
                 align=PP_ALIGN.LEFT)

    # Description
    add_text_box(slide,
                 "A branded desktop application that scrapes South African government "
                 "tender listings from etenders.gov.za, organises them into structured "
                 "Excel outputs, and keeps the team's equation tracker up to date — "
                 "all with a single click.",
                 left=Inches(0.7), top=Inches(4.88),
                 width=Inches(9.0), height=Inches(1.4),
                 font_size=13, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

    # Right panel text
    add_text_box(slide, "Beyond\nTechnology",
                 left=SLIDE_W - Inches(3.2), top=Inches(2.8),
                 width=Inches(3.0), height=Inches(1.2),
                 font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(slide, "Amidel (Pty) Ltd",
                 left=SLIDE_W - Inches(3.2), top=Inches(4.1),
                 width=Inches(3.0), height=Inches(0.4),
                 font_size=13, color=ORANGE, italic=True, align=PP_ALIGN.CENTER)

    # Bottom strip text
    add_text_box(slide, "Confidential  ·  Internal Use Only  ·  2026",
                 left=Inches(0.7), top=SLIDE_H - Inches(0.46),
                 width=Inches(9), height=Inches(0.4),
                 font_size=10, color=WHITE, align=PP_ALIGN.LEFT)


def slide_02_problem(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "The Problem", "Why manual tender tracking doesn't scale")
    slide_footer(slide, 2, TOTAL_SLIDES)

    top = content_area_top()
    pad = Inches(0.45)
    col_w = Inches(5.9)
    gap   = Inches(0.55)

    problems = [
        ("Time-consuming",
         ["etenders.gov.za must be visited daily",
          "Tenders span multiple pages — clicking through each one",
          "No export or download feature on the website"]),
        ("Error-prone",
         ["Manual copy-paste introduces mistakes",
          "Easy to miss tenders on busy days",
          "Duplicate entries go undetected"]),
        ("Hard to track trends",
         ["No historical view across batches",
          "Department-level counts require manual COUNTIF work",
          "Equation file maintained by hand every week"]),
        ("No structured output",
         ["Raw website data has no consistent format",
          "ICT vs RFQ classification done manually",
          "Briefing session details buried in expanded rows"]),
    ]

    card_h = Inches(1.55)
    for i, (heading, bullets) in enumerate(problems):
        col = i % 2
        row = i // 2
        left = pad + col * (col_w + gap)
        t = top + row * (card_h + Inches(0.18))
        bullet_card(slide, left, t, col_w, card_h, heading, bullets,
                    heading_size=14, bullet_size=12)


def slide_03_solution(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "The Solution", "One click. Structured output. Every time.")
    slide_footer(slide, 3, TOTAL_SLIDES)

    top = content_area_top()

    # Large central statement
    add_text_box(slide,
                 "The Amidel eTender Scraper is a branded desktop application "
                 "that automates the entire tender collection and reporting workflow "
                 "for the sales team — from website to Excel in one run.",
                 left=Inches(0.7), top=top,
                 width=Inches(11.9), height=Inches(0.95),
                 font_size=15, color=DARK_GRAY, align=PP_ALIGN.LEFT)

    # Three pillars
    pill_top = top + Inches(1.1)
    pill_w   = Inches(3.6)
    pill_h   = Inches(3.8)
    gap      = Inches(0.42)

    pillars = [
        ("Automate",
         "Selenium drives Chrome to visit each tender page, expand every row, "
         "and extract all 23 data fields — no manual steps.",
         ["Scrapes multiple days in sequence",
          "Handles pagination automatically",
          "Retry logic for network hiccups",
          "Duplicate detection built-in"]),
        ("Aggregate",
         "All daily files are merged into a single end-product workbook with "
         "a Tender Data tab and a Final tab containing COUNTIF formulas.",
         ["Per-day batches saved separately",
          "Combined RFQ & ICT Checker file",
          "Tender Summary per department",
          "Clickable hyperlinks in Excel"]),
        ("Act",
         "The equation tracker is updated automatically each run — "
         "new batch inserted, oldest dropped, borders applied, file saved.",
         ["Up to 6 rolling batches tracked",
          "Equation file copied to batch folder",
          "Synced via OneDrive automatically",
          "Ready to present immediately"]),
    ]

    for i, (title, desc, bullets) in enumerate(pillars):
        left = Inches(0.55) + i * (pill_w + gap)

        # Card background
        add_rect(slide, left, pill_top, pill_w, pill_h,
                 fill_rgb=NAVY, line_rgb=None)
        # Orange top strip
        add_rect(slide, left, pill_top, pill_w, Inches(0.07), fill_rgb=ORANGE)

        # Title
        add_text_box(slide, title,
                     left=left + Inches(0.2), top=pill_top + Inches(0.15),
                     width=pill_w - Inches(0.3), height=Inches(0.5),
                     font_size=18, bold=True, color=WHITE)

        # Description
        add_text_box(slide, desc,
                     left=left + Inches(0.2), top=pill_top + Inches(0.68),
                     width=pill_w - Inches(0.3), height=Inches(1.1),
                     font_size=11, color=LIGHT_GRAY)

        # Bullets
        txBox = slide.shapes.add_textbox(
            left + Inches(0.2), pill_top + Inches(1.85),
            pill_w - Inches(0.3), Inches(1.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        for j, b in enumerate(bullets):
            if j == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = "✓  " + b
            run.font.size = Pt(11.5)
            run.font.color.rgb = ORANGE
            run.font.name = "Segoe UI"


def slide_04_features(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "Key Features", "Everything the sales team needs, nothing it doesn't")
    slide_footer(slide, 4, TOTAL_SLIDES)

    top = content_area_top()
    features = [
        ("GUI — no command line", [
            "Amidel navy & orange branded desktop window",
            "Batch type toggle: Thursday (Mon–Wed) or Monday (Thu–Sun)",
            "Calendar popup with auto-highlighted batch range",
            "Live log output streamed to the screen while scraping",
        ]),
        ("Intelligent batch picker", [
            "Click any date in the week — correct range selected automatically",
            "Previous / next week navigation arrows",
            "Batch label displayed in human-readable format",
            "Batch type suggested automatically based on today's day",
        ]),
        ("Structured Excel output", [
            "Per-day files in batches/ subfolder",
            "Combined end-product file with Tender Data + Final tabs",
            "Tender Summary — one tab per department with any tenders (RFQs first)",
            "Clickable hyperlinks to source documents",
        ]),
        ("Equation file management", [
            "New batch inserted at column B (left-most position)",
            "Older batches shifted right automatically",
            "Maximum 6 batches enforced — oldest removed",
            "File copied to batch's Display Equation/ folder",
        ]),
        ("Robust scraping", [
            "Retry logic for stale elements and network errors",
            "Automatic duplicate detection and filtering",
            "pageLoadWait tuned for jQuery DataTable initialisation",
            "25 tracked departments / organs of state",
        ]),
        ("Organised folder structure", [
            "data/(T) DD-DD Mon YYYY/ named per batch",
            "Three subfolders: batches/, end product/, Display Equation/",
            "Scraper log written to logs/scraper.log",
            "OneDrive sync keeps equation file up to date",
        ]),
    ]

    card_h = Inches(1.55)
    card_w = Inches(3.9)
    gap_x  = Inches(0.42)
    gap_y  = Inches(0.18)
    pad    = Inches(0.45)

    for i, (heading, bullets) in enumerate(features):
        col = i % 3
        row = i // 3
        left = pad + col * (card_w + gap_x)
        t    = top  + row * (card_h + gap_y)
        bullet_card(slide, left, t, card_w, card_h, heading, bullets,
                    heading_size=13.5, bullet_size=11.5)


def slide_05_workflow(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "How It Works", "Five steps from launch to ready-to-present output")
    slide_footer(slide, 5, TOTAL_SLIDES)

    steps = [
        ("1", "Pick Batch",
         "Select batch type (T/M) and click any date in the target week. "
         "The app calculates the correct Mon–Wed or Thu–Sun range."),
        ("2", "Run Scraper",
         "Click Run Scraper. Selenium opens Chrome and visits "
         "etenders.gov.za for each day in the range."),
        ("3", "Extract Data",
         "Each tender row is expanded, all 23 fields are read, "
         "and duplicates are filtered before saving."),
        ("4", "Build Outputs",
         "Per-day Excel files are saved, then merged into the "
         "RFQ_and_ICT_Checker workbook and Tender Summary."),
        ("5", "Update Equation",
         "The equation file is updated with the new batch counts, "
         "saved, and copied to the Display Equation/ folder."),
    ]

    step_w = Inches(2.2)
    step_h = Inches(3.6)
    gap    = Inches(0.16)
    total  = len(steps) * step_w + (len(steps) - 1) * gap
    start  = (SLIDE_W - total) / 2
    top    = content_area_top() + Inches(0.3)

    for i, (num, title, desc) in enumerate(steps):
        left = start + i * (step_w + gap)

        # Card
        add_rect(slide, left, top, step_w, step_h, fill_rgb=NAVY)
        add_rect(slide, left, top, step_w, Inches(0.06), fill_rgb=ORANGE)

        # Number circle (simulated with small rect + big number)
        add_text_box(slide, num,
                     left=left, top=top + Inches(0.18),
                     width=step_w, height=Inches(0.8),
                     font_size=36, bold=True, color=ORANGE,
                     align=PP_ALIGN.CENTER)

        # Title
        add_text_box(slide, title,
                     left=left + Inches(0.1), top=top + Inches(1.0),
                     width=step_w - Inches(0.2), height=Inches(0.55),
                     font_size=14, bold=True, color=WHITE,
                     align=PP_ALIGN.CENTER)

        # Description
        add_text_box(slide, desc,
                     left=left + Inches(0.12), top=top + Inches(1.6),
                     width=step_w - Inches(0.24), height=Inches(1.8),
                     font_size=11, color=LIGHT_GRAY,
                     align=PP_ALIGN.CENTER)

        # Arrow between steps
        if i < len(steps) - 1:
            arrow_left = left + step_w + gap * 0.1
            arrow_top  = top + step_h / 2 - Inches(0.15)
            add_text_box(slide, "▶",
                         left=arrow_left, top=arrow_top,
                         width=gap * 0.8, height=Inches(0.3),
                         font_size=11, color=ORANGE, align=PP_ALIGN.CENTER)


def slide_06_outputs(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "Output Structure", "Everything you need, exactly where you expect it")
    slide_footer(slide, 6, TOTAL_SLIDES)

    top = content_area_top()

    # Folder tree on the left
    tree_left = Inches(0.5)
    tree_top  = top + Inches(0.1)
    tree_w    = Inches(5.6)
    tree_h    = Inches(5.3)

    add_rect(slide, tree_left, tree_top, tree_w, tree_h,
             fill_rgb=RGBColor(0x1A, 0x1A, 0x2E))

    tree_text = (
        "data/\n"
        "└── (T) 19-21 May 2026/\n"
        "    ├── batches/\n"
        "    │   ├── tenders_2026_05_19.xlsx\n"
        "    │   ├── tenders_2026_05_20.xlsx\n"
        "    │   └── tenders_2026_05_21.xlsx\n"
        "    ├── end product/\n"
        "    │   └── RFQ_and_ICT_Checker\n"
        "    │       _(19 - 21 May).xlsx\n"
        "    ├── Display Equation/\n"
        "    │   └── RFQ_and_ICT_Equation.xlsx\n"
        "    └── Tender Summary.xlsx\n"
        "\n"
        "logs/\n"
        "└── scraper.log"
    )

    add_text_box(slide, tree_text,
                 left=tree_left + Inches(0.2), top=tree_top + Inches(0.2),
                 width=tree_w - Inches(0.3), height=tree_h - Inches(0.3),
                 font_size=12, color=RGBColor(0xC8, 0xD0, 0xE0),
                 font_name="Consolas")

    # Right side — file descriptions
    desc_left = tree_left + tree_w + Inches(0.45)
    desc_w    = SLIDE_W - desc_left - Inches(0.4)

    files = [
        ("tenders_YYYY_MM_DD.xlsx",
         "Raw tender data for a single day — one row per tender, "
         "all 23 columns, saved immediately after each day is scraped."),
        ("RFQ_and_ICT_Checker_(DD - DD Month).xlsx",
         "Combined workbook: Tender Data tab (all tenders) + Final tab "
         "(COUNTIF formulas matching the equation file's source list)."),
        ("Tender Summary.xlsx",
         "One tab per tracked department that has any tenders. "
         "RFQs sort to the top; all other types sorted by closing date below."),
        ("RFQ_and_ICT_Equation.xlsx  [Display Equation/]",
         "Rolling tracker of NNT / ICT / RFQ counts per source. "
         "Newest batch always at column B; maximum 6 batches retained."),
        ("scraper.log",
         "Full structured log of every run — useful for diagnosing "
         "0-tender results or network / ChromeDriver failures."),
    ]

    card_h = Inches(0.94)
    gap    = Inches(0.12)
    for i, (fname, fdesc) in enumerate(files):
        t = tree_top + i * (card_h + gap)
        add_rect(slide, desc_left, t, desc_w, card_h,
                 fill_rgb=OFF_WHITE, line_rgb=LIGHT_GRAY, line_width_pt=0.5)
        add_rect(slide, desc_left, t, Inches(0.05), card_h, fill_rgb=ORANGE)

        add_text_box(slide, fname,
                     left=desc_left + Inches(0.15), top=t + Inches(0.06),
                     width=desc_w - Inches(0.2), height=Inches(0.35),
                     font_size=11, bold=True, color=NAVY,
                     font_name="Consolas")
        add_text_box(slide, fdesc,
                     left=desc_left + Inches(0.15), top=t + Inches(0.4),
                     width=desc_w - Inches(0.2), height=Inches(0.5),
                     font_size=10.5, color=DARK_GRAY)


def slide_07_data_fields(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "Data Captured", "23 fields extracted per tender")
    slide_footer(slide, 7, TOTAL_SLIDES)

    top = content_area_top()

    fields = [
        ("REPORT_DATE",                "Date the scraper was run"),
        ("RECORD_ID",                  "Auto-assigned ID (highest = first scraped)"),
        ("TENDER_ID",                  "Official tender number"),
        ("PUBLICATION_DATE",           "When the tender was published"),
        ("CLOSING_DATE / TIME",        "Submission deadline"),
        ("TENDER_TYPE",                "e.g. RFQ, Open Tender, etc."),
        ("TENDER_DESCRIPTION",         "Full tender title / description"),
        ("TENDER_SOURCE",              "Always ETENDERS.GOV.ZA"),
        ("DEPARTMENT",                 "Organ of state advertising the tender"),
        ("PROVINCE",                   "Province of the department"),
        ("ESUBMISSION",                "Whether e-submission is available"),
        ("CATEGORY",                   "Tender category (e.g. ICT, Construction)"),
        ("BRIEFING_DATE / VENUE",      "Briefing session details if applicable"),
        ("COMPULSORY_BRIEFING",        "Whether attendance is mandatory"),
        ("LINK",                       "Clickable hyperlink to tender document"),
        ("SOE",                        "State-owned enterprise flag"),
        ("COST_OF_SALES_ESTIMATE",     "Estimated cost / value"),
        ("CAPABILITY_AVAILABLE/GROUP", "Capability classification"),
        ("REQUIREMENTS",               "Any special requirements listed"),
    ]

    col_w  = Inches(5.7)
    row_h  = Inches(0.3)
    gap    = Inches(0.55)
    pad    = Inches(0.45)

    half = (len(fields) + 1) // 2
    for i, (field, desc) in enumerate(fields):
        col = i // half
        row = i % half
        left = pad + col * (col_w + gap)
        t    = top  + row * row_h

        # field name chip
        add_rect(slide, left, t + Inches(0.03), Inches(2.5), Inches(0.24),
                 fill_rgb=NAVY)
        add_text_box(slide, field,
                     left=left + Inches(0.05), top=t + Inches(0.03),
                     width=Inches(2.4), height=Inches(0.24),
                     font_size=9, bold=True, color=WHITE,
                     font_name="Consolas", align=PP_ALIGN.LEFT)
        add_text_box(slide, desc,
                     left=left + Inches(2.6), top=t + Inches(0.03),
                     width=col_w - Inches(2.65), height=Inches(0.24),
                     font_size=9.5, color=DARK_GRAY)


def slide_08_departments(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "Tracked Departments", "25 organs of state monitored each batch")
    slide_footer(slide, 8, TOTAL_SLIDES)

    top = content_area_top()

    depts = [
        "eTenders (All)", "Eskom", "SITA", "Transnet", "JPC",
        "Matatiele LM", "Ntabankulu LM", "Umzimvubu LM",
        "Winnie Mandela LM", "Mnquma LM", "Great Kei LM",
        "Amahlathi LM", "Raymond Mhlaba LM", "JOSHCO",
        "GPL", "RAF", "MM Trading Company",
        "Nelson Mandela Bay MM", "Buffalo City MM",
        "EC DPW", "DEL", "SIU", "GDoH", "DIRCO", "CP JHB / W JHB",
    ]

    # Intro text
    add_text_box(slide,
                 "For each batch the Final tab and Tender Summary produce counts of NNT, ICT, "
                 "and RFQ tenders broken down by the following departments. "
                 "New departments can be added by editing BatchProcessor.py.",
                 left=Inches(0.5), top=top,
                 width=Inches(12.3), height=Inches(0.65),
                 font_size=13, color=DARK_GRAY)

    # Chip grid
    chip_w = Inches(2.45)
    chip_h = Inches(0.42)
    gap_x  = Inches(0.18)
    gap_y  = Inches(0.14)
    cols   = 5
    pad    = Inches(0.5)
    start_top = top + Inches(0.8)

    for i, dept in enumerate(depts):
        col = i % cols
        row = i // cols
        left = pad + col * (chip_w + gap_x)
        t    = start_top + row * (chip_h + gap_y)

        # First chip is "eTenders (All)" — highlight it differently
        fill = NAVY if i == 0 else OFF_WHITE
        txt_col = WHITE if i == 0 else NAVY

        add_rect(slide, left, t, chip_w, chip_h,
                 fill_rgb=fill, line_rgb=NAVY_LIGHT, line_width_pt=0.6)
        if i > 0:
            add_rect(slide, left, t, Inches(0.04), chip_h, fill_rgb=ORANGE)

        add_text_box(slide, dept,
                     left=left + (Inches(0.08) if i > 0 else 0),
                     top=t + Inches(0.07),
                     width=chip_w - Inches(0.08), height=chip_h - Inches(0.08),
                     font_size=11, bold=(i == 0), color=txt_col,
                     align=PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT)


def slide_09_tech(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy_header_band(slide, "Technical Stack", "Built on proven, open-source Python libraries")
    slide_footer(slide, 9, TOTAL_SLIDES)

    top = content_area_top()

    stack = [
        ("Python 3", "Core language",
         "Clean, maintainable codebase with no compiled dependencies."),
        ("Tkinter + tkcalendar", "GUI framework",
         "Native desktop window with branded colours, calendar popup, and live log panel."),
        ("Selenium + ChromeDriver", "Web automation",
         "Drives Chrome headlessly or visibly. Automatically manages ChromeDriver version."),
        ("pandas", "Data processing",
         "DataFrame operations for deduplication, merging, and column transformations."),
        ("openpyxl", "Excel generation",
         "Writes formatted .xlsx files with styled headers, borders, and hyperlinks."),
        ("OneDrive", "File sync",
         "Equation file lives in OneDrive — updates are automatically available to all team members."),
    ]

    card_w = Inches(3.85)
    card_h = Inches(2.2)
    gap    = Inches(0.3)
    pad    = Inches(0.42)

    for i, (name, role, desc) in enumerate(stack):
        col = i % 3
        row = i // 3
        left = pad + col * (card_w + gap)
        t    = top  + row * (card_h + Inches(0.2))

        add_rect(slide, left, t, card_w, card_h,
                 fill_rgb=OFF_WHITE, line_rgb=LIGHT_GRAY, line_width_pt=0.75)
        add_rect(slide, left, t, Inches(0.06), card_h, fill_rgb=ORANGE)

        add_text_box(slide, name,
                     left=left + Inches(0.18), top=t + Inches(0.14),
                     width=card_w - Inches(0.25), height=Inches(0.45),
                     font_size=17, bold=True, color=NAVY)
        add_text_box(slide, role,
                     left=left + Inches(0.18), top=t + Inches(0.58),
                     width=card_w - Inches(0.25), height=Inches(0.3),
                     font_size=11, bold=True, color=ORANGE, italic=True)
        add_text_box(slide, desc,
                     left=left + Inches(0.18), top=t + Inches(0.9),
                     width=card_w - Inches(0.25), height=Inches(1.1),
                     font_size=12, color=DARK_GRAY)


def slide_10_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Full navy background
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=NAVY)
    add_rect(slide, 0, SLIDE_H - Inches(0.5), SLIDE_W, Inches(0.5), fill_rgb=ORANGE)
    add_rect(slide, SLIDE_W - Inches(3.4), 0, Inches(3.4), SLIDE_H - Inches(0.5),
             fill_rgb=NAVY_LIGHT)
    add_rect(slide, SLIDE_W - Inches(3.4), 0, Inches(0.055), SLIDE_H - Inches(0.5),
             fill_rgb=ORANGE)

    # Heading
    add_text_box(slide, "Summary",
                 left=Inches(0.7), top=Inches(0.9),
                 width=Inches(8.5), height=Inches(0.7),
                 font_size=40, bold=True, color=WHITE)
    add_text_box(slide, "What the Amidel eTender Scraper delivers",
                 left=Inches(0.7), top=Inches(1.6),
                 width=Inches(8.5), height=Inches(0.4),
                 font_size=15, color=ORANGE, italic=True)

    points = [
        "Saves hours of manual tender checking every batch week",
        "Produces consistently formatted, audit-ready Excel files",
        "Tracks NNT / ICT / RFQ counts across 6 rolling batches",
        "Monitors 25 departments and organs of state automatically",
        "One-click operation — no technical knowledge required",
        "Runs on any Windows machine with Chrome installed",
    ]

    for i, pt in enumerate(points):
        add_text_box(slide, "✓  " + pt,
                     left=Inches(0.85), top=Inches(2.25) + i * Inches(0.62),
                     width=Inches(8.5), height=Inches(0.55),
                     font_size=14, color=WHITE)

    # Right panel
    add_text_box(slide, "Get\nStarted",
                 left=SLIDE_W - Inches(3.2), top=Inches(1.5),
                 width=Inches(3.0), height=Inches(1.1),
                 font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    steps = [
        "pip install -r requirements.txt",
        "python main.py",
        "Pick batch type & week",
        "Click Run Scraper",
    ]
    txBox = slide.shapes.add_textbox(
        SLIDE_W - Inches(3.15), Inches(2.7),
        Inches(2.9), Inches(3.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    for j, s in enumerate(steps):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = f"{j + 1}.  {s}"
        run.font.size = Pt(11)
        run.font.color.rgb = LIGHT_GRAY if j > 0 else ORANGE
        run.font.name = "Consolas" if j < 2 else "Segoe UI"
        run.font.bold = j == 0

    add_text_box(slide, "Amidel (Pty) Ltd  ·  2026",
                 left=Inches(0.7), top=SLIDE_H - Inches(0.46),
                 width=Inches(9), height=Inches(0.4),
                 font_size=10, color=WHITE)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    prs = new_prs()

    slide_01_title(prs)
    slide_02_problem(prs)
    slide_03_solution(prs)
    slide_04_features(prs)
    slide_05_workflow(prs)
    slide_06_outputs(prs)
    slide_07_data_fields(prs)
    slide_08_departments(prs)
    slide_09_tech(prs)
    slide_10_summary(prs)

    out = "Amidel_eTender_Scraper_Presentation.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
