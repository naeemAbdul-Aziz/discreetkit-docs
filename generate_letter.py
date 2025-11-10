import os
from datetime import datetime
from typing import List
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

# Define color palette from user
COLOR_CYAN_TURQUOISE = colors.HexColor("#187f76")
COLOR_INDIGO = colors.HexColor("#1e3a5f")
COLOR_AZTEC_GOLD = colors.HexColor("#c48c52")
COLOR_METALLIC_YELLOW = colors.HexColor("#ffce07")
COLOR_LIGHT_SILVER = colors.HexColor("#d7d9db")

# Define asset paths
ASSETS_DIR = os.path.join(os.path.dirname(__file__), 'assets')
FONT_REGULAR_PATH = os.path.join(ASSETS_DIR, 'Satoshi-Regular.ttf')
FONT_BOLD_PATH = os.path.join(ASSETS_DIR, 'Satoshi-Bold.ttf')
LOGO_PATH = os.path.join(ASSETS_DIR, 'logo.png')

# --- Font Registration ---

# Ensure font files exist before proceeding
if not os.path.exists(FONT_REGULAR_PATH) or not os.path.exists(FONT_BOLD_PATH):
    print(f"Error: Font files not found.")
    print(f"Please place 'Satoshi-Regular.ttf' and 'Satoshi-Bold.ttf' in the '{ASSETS_DIR}' directory.")
    # In a real app, you'd raise an Exception here
    # exit() 
else:
    pdfmetrics.registerFont(TTFont('Satoshi', FONT_REGULAR_PATH))
    pdfmetrics.registerFont(TTFont('Satoshi-Bold', FONT_BOLD_PATH))

 # Use registered Satoshi fonts when available, otherwise fall back to built-in fonts
if 'Satoshi' in pdfmetrics.getRegisteredFontNames():
    FONT_REGULAR_NAME = 'Satoshi'
else:
    FONT_REGULAR_NAME = 'Helvetica'

if 'Satoshi-Bold' in pdfmetrics.getRegisteredFontNames():
    FONT_BOLD_NAME = 'Satoshi-Bold'
else:
    FONT_BOLD_NAME = 'Helvetica-Bold'


def draw_header_footer(canv, doc):
    """Draw header and footer on each page using the canvas and doc.

    This function is used as the onFirstPage/onLaterPages callback for
    SimpleDocTemplate.build(...) so header/footer are drawn independently
    from the story flowables and won't overlap the document content.
    """
    canv.saveState()

    width, height = doc.pagesize

    # --- Header ---
    header_height = 40 * mm

    # Cyan sidebar at top-left (short bar)
    canv.setFillColor(COLOR_CYAN_TURQUOISE)
    canv.rect(0, height - header_height, 10 * mm, header_height, stroke=0, fill=1)

    # Logo (if present) placed near top-left inside page margins
    draw_w = 0
    draw_h = 0
    if os.path.exists(LOGO_PATH):
        try:
            logo = ImageReader(LOGO_PATH)
            logo_w, logo_h = logo.getSize()
            aspect = logo_h / float(logo_w)
            draw_w = 1.25 * inch
            draw_h = draw_w * aspect
            canv.drawImage(logo, 20 * mm, height - (1 * inch) - draw_h, width=draw_w, height=draw_h, mask='auto')
        except Exception:
            draw_w = 0
            draw_h = 0

    # Company name to the right of logo
    try:
        canv.setFont(FONT_REGULAR_NAME, 9)
    except Exception:
        canv.setFont('Helvetica', 9)
    canv.setFillColor(COLOR_INDIGO)
    canv.drawString(20 * mm + draw_w + (5 * mm), height - (1 * inch) - (draw_h / 2) + 5, "ACCESS DISCREETKIT LTD")

    # Header separator line (below header area)
    header_bottom = height - header_height - (3 * mm)
    canv.setStrokeColor(COLOR_LIGHT_SILVER)
    canv.setLineWidth(0.5)
    canv.line(20 * mm, header_bottom, width - (20 * mm), header_bottom)

    # --- Footer ---
    footer_y = doc.bottomMargin - (6 * mm)
    if footer_y < 12 * mm:
        footer_y = 12 * mm

    styles = get_letter_styles()
    footer_style = styles['Footer']

    # Column 1: Address (left)
    address_text = """
        <b>Address</b><br/>
        House No. 57, Kofi Annan East Avenue,<br/>
        Madina, Accra, Ghana
    """
    p_address = Paragraph(address_text, footer_style)
    w_addr, h_addr = p_address.wrap(width / 3, doc.bottomMargin)
    p_address.drawOn(canv, 20 * mm, footer_y)

    # Column 2: Contact (center)
    contact_text = """
        <b>Contact</b><br/>
        Email: discreetkit@gmail.com<br/>
        Phone: +233 20 300 1107
    """
    p_contact = Paragraph(contact_text, footer_style)
    w_con, h_con = p_contact.wrap(width / 3, doc.bottomMargin)
    p_contact.drawOn(canv, (width / 2) - (w_con / 2), footer_y)

    # Column 3: Social (right)
    social_text = """
        <b>Follow Us</b><br/>
        Twitter: @discreetkit<br/>
        LinkedIn: /company/discreetkit
    """
    p_social = Paragraph(social_text, footer_style)
    w_soc, h_soc = p_social.wrap(width / 3, doc.bottomMargin)
    p_social.drawOn(canv, width - (20 * mm) - w_soc, footer_y)

    canv.restoreState()


# --- Styles ---
def get_letter_styles():
    """Returns a stylesheet configured for the letter.

    This updates existing named styles when present to avoid duplicate
    KeyError from reportlab's stylesheet when called multiple times.
    """
    styles = getSampleStyleSheet()

    # Update base Normal style
    styles['Normal'].fontName = FONT_REGULAR_NAME
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 14
    styles['Normal'].textColor = COLOR_INDIGO

    # Recipient
    if 'Recipient' in styles:
        styles['Recipient'].spaceBefore = 10
    else:
        styles.add(ParagraphStyle(name='Recipient', parent=styles['Normal'], spaceBefore=10))

    # Date
    if 'Date' in styles:
        styles['Date'].alignment = TA_RIGHT
    else:
        styles.add(ParagraphStyle(name='Date', parent=styles['Normal'], alignment=TA_RIGHT))

    # Heading1
    if 'Heading1' in styles:
        styles['Heading1'].fontName = FONT_BOLD_NAME
        styles['Heading1'].fontSize = 18
        styles['Heading1'].textColor = COLOR_INDIGO
        styles['Heading1'].spaceBefore = 12
        styles['Heading1'].spaceAfter = 6
    else:
        styles.add(ParagraphStyle(name='Heading1', parent=styles['Normal'], fontName=FONT_BOLD_NAME, fontSize=18, textColor=COLOR_INDIGO, spaceBefore=12, spaceAfter=6))

    # Body
    if 'Body' in styles:
        styles['Body'].spaceAfter = 12
    else:
        styles.add(ParagraphStyle(name='Body', parent=styles['Normal'], spaceAfter=12))

    # Footer
    if 'Footer' in styles:
        styles['Footer'].fontName = FONT_REGULAR_NAME
        styles['Footer'].fontSize = 8
        styles['Footer'].leading = 10
        styles['Footer'].textColor = COLOR_INDIGO
    else:
        styles.add(ParagraphStyle(name='Footer', parent=styles['Normal'], fontName=FONT_REGULAR_NAME, fontSize=8, leading=10, textColor=COLOR_INDIGO))

    return styles


# --- Main Document Generation ---

def generate_document(filename, content):
    """
    Generates the PDF letter.
    
    :param filename: The output PDF file name (e.g., "proposal.pdf")
    :param content: A dictionary containing the letter's text components.
    """
    
    # --- Page Setup ---
    # We use a large top margin for the logo/header area
    # and a large left margin for the cyan bar.
    # FIXED: Increased topMargin to 2.0 inch for taller header
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        topMargin=2.0 * inch,
        bottomMargin=1.25 * inch,
        leftMargin=20 * mm, # Main content margin
        rightMargin=20 * mm
    )
    styles = get_letter_styles()
    # story holds the document flowables (paragraphs, spacers, etc.)
    story = []

    # --- Build Story ---

    # 1. Date
    story.append(Paragraph(content.get('date', ''), styles['Date']))
    story.append(Spacer(1, 0.25 * inch))

    # 2. Recipient Info
    recipient_lines = content.get('recipient', [])
    for line in recipient_lines:
        story.append(Paragraph(line, styles['Recipient']))
    
    story.append(Spacer(1, 0.25 * inch))

    # 3. Document Title
    story.append(Paragraph(content.get('title', ''), styles['Heading1']))
    story.append(Spacer(1, 0.1 * inch))

    # 4. Salutation
    story.append(Paragraph(content.get('salutation', ''), styles['Normal']))
    story.append(Spacer(1, 0.1 * inch))

    # 5. Body
    body_paragraphs = content.get('body', [])
    for para in body_paragraphs:
        story.append(Paragraph(para, styles['Body']))

    # 6. Closing
    story.append(Paragraph(content.get('closing', ''), styles['Normal']))
    story.append(Spacer(1, 0.25 * inch)) # Space for signature

    # 7. Signature
    signature_lines = content.get('signature', [])
    for line in signature_lines:
        story.append(Paragraph(line, styles['Normal']))
        
    # --- Generate PDF ---
    try:
        # Draw header/footer via callbacks so they don't overlap story content
        doc.build(story, onFirstPage=draw_header_footer, onLaterPages=draw_header_footer)
        print(f"Successfully generated '{filename}'")
    except Exception as e:
        print(f"Error generating PDF: {e}")

# --- Example Usage ---

if __name__ == "__main__":
    
    # Define the dynamic content for the letter
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
            "Thank you for considering our proposal. We look forward to the possibility of working together."
        ],
        
        "closing": "Sincerely,",
        
        "signature": [
            "Jane Smith",
            "Director of Business Development",
            "Access DiscreetKit Ltd"
        ]
    }
    
    # Generate the document
    generate_document("strategic_proposal.pdf", letter_content)