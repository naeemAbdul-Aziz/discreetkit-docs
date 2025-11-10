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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Flowable

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

# --- Custom Styles ---

def get_letter_styles():
    """Returns a stylesheet library for the document."""
    styles = getSampleStyleSheet()

    # Update the existing Normal style rather than re-adding it (avoids duplicate-name errors)
    styles['Normal'].fontName = FONT_REGULAR_NAME
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 14
    styles['Normal'].textColor = COLOR_INDIGO

    # Helper: add a style if it doesn't exist, otherwise update attributes
    def add_or_update_style(stylesheet, name, **style_kwargs):
        if name in stylesheet:
            st = stylesheet[name]
            # apply attributes where possible
            for k, v in style_kwargs.items():
                try:
                    setattr(st, k, v)
                except Exception:
                    # ignore unknown attrs
                    pass
        else:
            stylesheet.add(ParagraphStyle(name=name, **style_kwargs))

    add_or_update_style(styles, 'Recipient', parent=styles['Normal'], spaceBefore=10)
    add_or_update_style(styles, 'Date', parent=styles['Normal'], alignment=TA_RIGHT)
    add_or_update_style(styles, 'Heading1', parent=styles['Normal'], fontName=FONT_BOLD_NAME,
                        fontSize=18, textColor=COLOR_INDIGO, spaceBefore=12, spaceAfter=6)
    add_or_update_style(styles, 'Body', parent=styles['Normal'], spaceAfter=12)
    add_or_update_style(styles, 'Footer', parent=styles['Normal'], fontName=FONT_REGULAR_NAME,
                        fontSize=8, leading=10, textColor=COLOR_INDIGO)

    return styles

# --- Header/Footer Drawing Functions ---

class HeaderFooter(Flowable):
    """
    A flowable to draw header and footer content. This is used to
    ensure elements are drawn *under* the main content flow.
    """
    def __init__(self):
        Flowable.__init__(self)
        self.width, self.height = A4
        # Check for logo
        self.logo = None
        # default draw sizes in case logo missing
        self._draw_w = 0
        self._draw_h = 0
        if os.path.exists(LOGO_PATH):
            self.logo = ImageReader(LOGO_PATH)
        else:
            print(f"Warning: Logo file not found at '{LOGO_PATH}'. Header will not include logo.")

    def wrap(self, *args):
        # Doesn't take up space in the story
        return (0, 0)

    def draw(self):
        self.canv.saveState()
        
        # --- Header Elements ---
        
        # 1. Cyan sidebar (thinner, as recommended)
        self.canv.setFillColor(COLOR_CYAN_TURQUOISE)
        self.canv.rect(0, 0, 10 * mm, self.height, stroke=0, fill=1)
        
        # 2. Logo
        if self.logo:
            # Draw logo at 1 inch from top, 1 inch from left (plus sidebar)
            # Assuming logo is 1.5 inch wide, 0.5 inch high. Adjust as needed.
            logo_w, logo_h = self.logo.getSize()
            aspect = logo_h / float(logo_w)
            draw_w = 1.25 * inch
            draw_h = draw_w * aspect
            # store for use later (company name positioning)
            self._draw_w = draw_w
            self._draw_h = draw_h
            # Position inside left margin (set in SimpleDocTemplate)
            self.canv.drawImage(self.logo, 20 * mm, self.height - (1 * inch) - draw_h, 
                                width=draw_w, height=draw_h, mask='auto')

        # 3. Company Name (if not part of logo)
        # Use the resolved font names (fallbacks applied earlier)
        try:
            self.canv.setFont(FONT_REGULAR_NAME, 9)
        except Exception:
            # final fallback to a built-in font name
            self.canv.setFont('Helvetica', 9)
        self.canv.setFillColor(COLOR_INDIGO)
        # Use stored draw size values (0 if logo missing)
        self.canv.drawString(20 * mm + self._draw_w + (5 * mm), self.height - (1 * inch) - (self._draw_h / 2) + 5, 
                             "ACCESS DISCREETKIT LTD")

        # --- Footer Elements ---
        footer_y = 1 * inch
        
        # 1. Horizontal Rule
        self.canv.setStrokeColor(COLOR_LIGHT_SILVER)
        self.canv.setLineWidth(0.5)
        self.canv.line(20 * mm, footer_y, self.width - (20 * mm), footer_y)

        # 2. Footer Content (Two-Column)
        styles = get_letter_styles()
        footer_style = styles['Footer']
        
        # Column 1: Address
        address_text = """
            <b>Address</b><br/>
            House No. 57, Kofi Annan East Avenue,<br/>
            Madina, Accra, Ghana
        """
        p_address = Paragraph(address_text, footer_style)
        w, h = p_address.wrapOn(self.canv, self.width / 2 - (25 * mm), footer_y)
        p_address.drawOn(self.canv, 20 * mm, footer_y - h - (3 * mm))
        
        # Column 2: Contact
        contact_text = """
            <b>Contact</b><br/>
            Email: discreetkit@gmail.com<br/>
            Phone: +233 20 300 1107
        """
        p_contact = Paragraph(contact_text, footer_style)
        w, h = p_contact.wrapOn(self.canv, self.width / 3, footer_y)
        p_contact.drawOn(self.canv, self.width / 2, footer_y - h - (3 * mm))

        # Column 3: Social (if needed)
        social_text = """
            <b>Follow Us</b><br/>
            Twitter: @discreetkit<br/>
            LinkedIn: /company/discreetkit
        """
        p_social = Paragraph(social_text, footer_style)
        w, h = p_social.wrapOn(self.canv, self.width / 3, footer_y)
        p_social.drawOn(self.canv, self.width - (20 * mm) - w, footer_y - h - (3 * mm))

        self.canv.restoreState()


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
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        topMargin=1.5 * inch,
        bottomMargin=1.25 * inch,
        leftMargin=20 * mm, # Main content margin
        rightMargin=20 * mm
    )
    styles = get_letter_styles()
    story: List[Flowable] = [HeaderFooter()]  # Add the header/footer flowable first
    story = [HeaderFooter()] # Add the header/footer flowable first

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
        doc.build(story)
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