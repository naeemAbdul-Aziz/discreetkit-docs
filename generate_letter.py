import os
from datetime import datetime
from typing import List, Dict
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak

# --- Configuration ---
COLOR_CYAN_TURQUOISE = colors.HexColor("#187f76")
COLOR_INDIGO = colors.HexColor("#1e3a5f")
COLOR_AZTEC_GOLD = colors.HexColor("#c48c52")
COLOR_METALLIC_YELLOW = colors.HexColor("#ffce07")
COLOR_LIGHT_SILVER = colors.HexColor("#d7d9db")

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
FONT_REGULAR_PATH = os.path.join(ASSETS_DIR, "Satoshi-Regular.ttf")
FONT_BOLD_PATH = os.path.join(ASSETS_DIR, "Satoshi-Bold.ttf")
LOGO_PATH = os.path.join(ASSETS_DIR, "logo.png")

# --- Font Registration with fallbacks ---
if os.path.exists(FONT_REGULAR_PATH):
    try:
        pdfmetrics.registerFont(TTFont("Satoshi", FONT_REGULAR_PATH))
    except Exception:
        pass
if os.path.exists(FONT_BOLD_PATH):
    try:
        pdfmetrics.registerFont(TTFont("Satoshi-Bold", FONT_BOLD_PATH))
    except Exception:
        pass

# resolved font names (fall back to built-in fonts)
REGISTERED = set(pdfmetrics.getRegisteredFontNames())
FONT_REGULAR_NAME = "Satoshi" if "Satoshi" in REGISTERED else "Helvetica"
FONT_BOLD_NAME = "Satoshi-Bold" if "Satoshi-Bold" in REGISTERED else "Helvetica-Bold"


def add_or_update_style(styles, name: str, parent_name: str = "Normal", **props):
    """Add or update a ParagraphStyle in a stylesheet."""
    if name in styles:
        s = styles[name]
        for k, v in props.items():
            setattr(s, k, v)
    else:
        parent = styles[parent_name] if parent_name in styles else styles["Normal"]
        ps = ParagraphStyle(name=name, parent=parent, **props)
        styles.add(ps)


def get_letter_styles():
    """Return a stylesheet and update existing named styles if present."""
    styles = getSampleStyleSheet()

    # Base Normal style
    styles["Normal"].fontName = FONT_REGULAR_NAME
    styles["Normal"].fontSize = 10
    styles["Normal"].leading = 14
    styles["Normal"].textColor = COLOR_INDIGO

    add_or_update_style(styles, "Recipient", spaceBefore=10)
    add_or_update_style(styles, "Date", alignment=TA_RIGHT)
    add_or_update_style(
        styles,
        "Heading1",
        fontName=FONT_BOLD_NAME,
        fontSize=18,
        textColor=COLOR_INDIGO,
        spaceBefore=12,
        spaceAfter=6,
    )
    add_or_update_style(styles, "Body", spaceAfter=12)
    add_or_update_style(
        styles,
        "Footer",
        fontName=FONT_REGULAR_NAME,
        fontSize=8,
        leading=10,
        textColor=COLOR_INDIGO,
    )

    return styles


def draw_header_footer(canv, doc):
    """Draw header and footer on every page.

    Footer columns are distributed evenly within the left/right margins and
    placed slightly lower to avoid overlapping content.
    """
    canv.saveState()
    width, height = doc.pagesize

    # --- Header ---
    header_height = 40 * mm

    # Cyan sidebar at top-left
    canv.setFillColor(COLOR_CYAN_TURQUOISE)
    canv.rect(0, height - header_height, 10 * mm, header_height, stroke=0, fill=1)

    # Logo
    draw_w = 0
    draw_h = 0
    if os.path.exists(LOGO_PATH):
        try:
            logo = ImageReader(LOGO_PATH)
            logo_w, logo_h = logo.getSize()
            aspect = logo_h / float(logo_w)
            draw_w = 1.25 * inch
            draw_h = draw_w * aspect
            canv.drawImage(logo, doc.leftMargin, height - (1 * inch) - draw_h, width=draw_w, height=draw_h, mask='auto')
        except Exception:
            draw_w = 0
            draw_h = 0

    # Company name
    try:
        canv.setFont(FONT_REGULAR_NAME, 9)
    except Exception:
        canv.setFont('Helvetica', 9)
    canv.setFillColor(COLOR_INDIGO)
    canv.drawString(doc.leftMargin + draw_w + (5 * mm), height - (1 * inch) - (draw_h / 2) + 5, "ACCESS DISCREETKIT LTD")

    # Header separator
    header_bottom = height - header_height - (3 * mm)
    canv.setStrokeColor(COLOR_LIGHT_SILVER)
    canv.setLineWidth(0.5)
    canv.line(doc.leftMargin, header_bottom, width - doc.rightMargin, header_bottom)

    # --- Footer ---
    left = doc.leftMargin
    right = width - doc.rightMargin
    usable_width = right - left

    # Move footer lower to avoid overlap
    footer_line_y = doc.bottomMargin - (6 * mm)
    if footer_line_y < 12 * mm:
        footer_line_y = 12 * mm

    canv.setStrokeColor(COLOR_LIGHT_SILVER)
    canv.setLineWidth(0.5)
    canv.line(left, footer_line_y, left + usable_width, footer_line_y)

    styles = get_letter_styles()
    footer_style = styles['Footer']

    cols = [
        "<b>Address</b><br/>House No. 57, Kofi Annan East Avenue,<br/>Madina, Accra, Ghana",
        "<b>Contact</b><br/>Email: discreetkit@gmail.com<br/>Phone: +233 20 300 1107",
        "<b>Follow Us</b><br/>Twitter: @discreetkit<br/>LinkedIn: /company/discreetkit",
    ]

    col_width = usable_width / 3.0
    x_positions = [left + i * col_width for i in range(3)]
    padding = 3 * mm

    for i, txt in enumerate(cols):
        p = Paragraph(txt, footer_style)
        w, h = p.wrap(col_width, doc.bottomMargin)
        draw_x = x_positions[i]
        draw_y = footer_line_y - padding - h
        if draw_y < 6 * mm:
            draw_y = 6 * mm
        p.drawOn(canv, draw_x, draw_y)

    canv.restoreState()


def generate_document(filename: str, content: dict):
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        topMargin=2.0 * inch,
        bottomMargin=1.25 * inch,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
    )

    styles = get_letter_styles()
    story = []

    # Small spacer so content doesn't touch header
    story.append(Spacer(1, 0.2 * inch))

    # Date
    story.append(Paragraph(str(content.get('date', '')), styles['Date']))
    story.append(Spacer(1, 0.25 * inch))

    # Recipient
    recipient = content.get('recipient', [])
    if isinstance(recipient, list):
        for line in recipient:
            story.append(Paragraph(str(line), styles['Recipient']))

    story.append(Spacer(1, 0.25 * inch))

    # Title
    story.append(Paragraph(str(content.get('title', '')), styles['Heading1']))
    story.append(Spacer(1, 0.1 * inch))

    # Salutation
    story.append(Paragraph(str(content.get('salutation', '')), styles['Normal']))
    story.append(Spacer(1, 0.1 * inch))

    # Body
    body = content.get('body', [])
    if isinstance(body, list):
        for para in body:
            story.append(Paragraph(str(para), styles['Body']))

    # Closing and signature
    story.append(Paragraph(str(content.get('closing', '')), styles['Normal']))
    story.append(Spacer(1, 0.25 * inch))
    signature = content.get('signature', [])
    if isinstance(signature, list):
        for line in signature:
            story.append(Paragraph(str(line), styles['Normal']))

    try:
        doc.build(story, onFirstPage=draw_header_footer, onLaterPages=draw_header_footer)
        print(f"Successfully generated '{filename}'")
    except Exception as e:
        print(f"Error generating PDF: {e}")


# --- Example Usage ---
if __name__ == "__main__":
    letter_content = {
        "date": datetime.now().strftime("%d %B %Y"),
        "recipient": [
            "Mr. John Doe",
            "Chief Executive Officer",
            "Innovate Corp.",
            "123 Innovation Drive, Accra, Ghana"
        ],
        "title": "Proposal for Strategic Partnership",
        "salutation": "Dear Mr. Doe,",
        "body": [
            "We are writing to propose a strategic partnership between Access DiscreetKit Ltd and Innovate Corp. Our analysis indicates that a collaboration could unlock significant value in the market. We have attached a detailed deck outlining the potential synergies, go-to-market strategy, and proposed financial arrangements.",
            "This is a second paragraph to demonstrate text wrapping and flow. It will continue on as long as necessary, respecting the margins we've set. When this paragraph becomes too long for the current page, it will automatically break and continue on a new page, which will also feature the same header and footer.",
            "Thank you for considering our proposal. We look forward to the possibility of working together.",
        ],
        "closing": "Sincerely,",
        "signature": [
            "Jane Smith",
            "Director of Business Development",
            "Access DiscreetKit Ltd",
        ],
    }

    generate_document("strategic_proposal.pdf", letter_content)