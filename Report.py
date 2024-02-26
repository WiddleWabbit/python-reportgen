import decimal

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import Paragraph


class ReportGenerator:

    colours = {
        "black": (0, 0, 0),
        "darkgrey": (0.1, 0.1, 0.1),
        "white": (1, 1, 1),
        "blue": (0, 0.539, 0.625),
        "purple": (0.34375, 0.296875, 0.4375)
    }

    page_num: int = 1

    def __init__(self, output_file=str, asset_folder=str):

        self.output_file = output_file
        self.c = canvas.Canvas(self.output_file, pagesize=A4)
        self.page_width, self.page_height = A4
        self.register_fonts(asset_folder)
        self.styles = self.get_styles()
        self.asset_folder = asset_folder

    @staticmethod
    def register_fonts(asset_folder):
        # Register Fonts
        pdfmetrics.registerFont(TTFont('OpenSans-Bold', asset_folder + "OpenSans-Bold.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-BoldItalic', asset_folder + "OpenSans-BoldItalic.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-ExtraBold', asset_folder + "OpenSans-ExtraBold.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-ExtraBoldItalic', asset_folder + "OpenSans-ExtraBoldItalic.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-Italic', asset_folder + "OpenSans-Italic.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-Light', asset_folder + "OpenSans-Light.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-LightItalic', asset_folder + "OpenSans-LightItalic.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-Medium', asset_folder + "OpenSans-Medium.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-MediumItalic', asset_folder + "OpenSans-MediumItalic.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-Regular', asset_folder + "OpenSans-Regular.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-SemiBold', asset_folder + "OpenSans-SemiBold.ttf"))
        pdfmetrics.registerFont(TTFont('OpenSans-SemiBoldItalic', asset_folder + "OpenSans-SemiBoldItalic.ttf"))

    def get_styles(self):
        styles = getSampleStyleSheet()
        # Create a custom style based on an existing style
        body_style = ParagraphStyle(
            name='SKN-Body',
            parent=styles['Normal'],
            fontName='OpenSans-Regular',
            wordWrap='default',
            textColor=self.colours["darkgrey"],
            fontSize=9,
            charSpace=0,
        )
        title_style = ParagraphStyle(
            name='SKN-Title',
            parent=styles['Normal'],
            fontName='OpenSans-ExtraBold',
            wordWrap='none',
            textColor=self.colours["blue"],
            fontSize=32,
            charSpace=0,
        )
        sub_title_style = ParagraphStyle(
            name='SKN-Sub-Title',
            parent=styles['Normal'],
            fontName='OpenSans-Bold',
            wordWrap='none',
            textColor=self.colours["purple"],
            fontSize=15,
            charSpace=0,
        )
        heading_style = ParagraphStyle(
            name='SKN-Heading',
            parent=styles['Normal'],
            fontName='OpenSans-Bold',
            wordWrap='none',
            textColor=self.colours["black"],
            fontSize=13.75,
            charSpace=0,
        )
        sub_heading_style = ParagraphStyle(
            name='SKN-Sub-Heading',
            parent=styles['Normal'],
            fontName='OpenSans-Regular',
            wordWrap='none',
            textColor=self.colours["black"],
            fontSize=11,
            charSpace=0,
        )
        coloured_heading_style = ParagraphStyle(
            name='SKN-Coloured-Heading',
            parent=styles['Normal'],
            fontName='OpenSans-ExtraBold',
            wordWrap='none',
            textColor=self.colours["blue"],
            fontSize=17,
            charSpace=0,
        )
        footer_blue_style = ParagraphStyle(
            name='SKN-Footer-Blue',
            parent=styles['Normal'],
            fontName='OpenSans-Light',
            wordWrap='none',
            textColor=self.colours["blue"],
            fontSize=10.5,
            charSpace=1,
        )
        footer_purple_style = ParagraphStyle(
            name='SKN-Footer-Purple',
            parent=styles['Normal'],
            fontName='OpenSans-Light',
            wordWrap='none',
            textColor=self.colours["purple"],
            fontSize=10.5,
            charSpace=1,
        )
        caption_style = ParagraphStyle(
            name='SKN-Caption',
            parent=styles['Normal'],
            fontName='OpenSans-Bold',
            wordWrap='none',
            textColor=self.colours["white"],
            fontSize=9,
            charSpace=0,
        )
        styles.add(body_style)
        styles.add(title_style)
        styles.add(sub_title_style)
        styles.add(heading_style)
        styles.add(sub_heading_style)
        styles.add(footer_blue_style)
        styles.add(footer_purple_style)
        styles.add(caption_style)
        styles.add(coloured_heading_style)

        return styles

    def add_header_footer(self):
        # Place Header & Footer Lines
        self.add_line(20, 190, 270, 270, 0.5, self.colours["purple"])
        self.add_line(20, 190, 20, 20, 0.5, self.colours["purple"])
        # Place the company logo
        self.add_img(20, 274, 46, 12.451, self.asset_folder + "SkinElementsLogo.jpg")
        # Place the Footer Text
        self.add_text(20, 12, str(self.page_num) + " |", "SKN-Footer-Blue")
        self.add_text(26, 12, " SKIN ELEMENTS LIMITED |", "SKN-Footer-Purple")
        self.add_text(79, 12, " FACE UV ANALYSIS ", "SKN-Footer-Blue")

    def add_rect(self, x, y, width, height, colour: tuple, opacity: float, fill: bool, stroke: bool):
        self.c.setFillColorRGB(*colour, alpha=opacity)
        self.c.rect(x * mm, y * mm, width * mm, height * mm, fill=fill, stroke=stroke)

    def add_line(self, x_start, x_end, y_start, y_end, width, colour):
        self.c.setLineWidth(width)
        self.c.setStrokeColorRGB(*colour)
        # Add the Line
        self.c.line(x_start * mm, y_start * mm, x_end * mm, y_end * mm)

    def add_text(self, x, y, text, style):
        style_info = self.styles[style]
        self.c.setFont(style_info.fontName, style_info.fontSize)
        self.c.setFillColorRGB(*style_info.textColor)
        char_spacing = f", charSpace={style_info.charSpace}"
        # Place the specified text at the specified location
        self.c.drawString(x * mm, y * mm, text, mode=0, charSpace=style_info.charSpace)

    def add_para(self, x, y, width, height, text, style):
        # Create a paragraph object containing the correct text in the right style
        para = Paragraph(text, self.styles[f"{style}"])
        # Wrap the paragraph in the bounding box defined here
        para.wrapOn(self.c, width * mm, height * mm)
        # Place the paragraph
        para.drawOn(self.c, x * mm, y * mm)

    def add_img(self, x, y, width, height, file: str):
        self.c.drawImage(file, x * mm, y * mm, width * mm, height * mm)

    def add_page(self):
        self.c.showPage()

    def inc_page_num(self):
        self.page_num = self.page_num + 1

    def set_page_num(self, num: int):
        self.page_num = num

    def save_report(self):
        self.c.save()
