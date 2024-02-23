import os
import re
import glob
import PyPDF2
import shutil
import math

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import Paragraph


# Requirements PyPDF2, Pillow, ReportLab


# Function to check if there are at least two JPG files and at least one PDF file in a folder
def check_files(folder_path):
    jpg_files = glob.glob(os.path.join(folder_path, '*.jpg'))
    pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))

    if len(jpg_files) >= 2 and len(pdf_files) >= 1:
        return True
    else:
        raise ValueError("Folder does not contain at least two JPG files and one PDF file.")


def resize_image(input_image_path, output_image_path, size):
    with Image.open(input_image_path) as image:
        resized_image = image.resize(size)
        resized_image = resized_image.crop((155, 0, 2564, 3508))
        resized_image.save(output_image_path)


def split_image(input_image_path, output_image_folder):
    # Open an image file
    with Image.open(input_image_path) as img:
        # Get the dimensions of the image
        width, height = img.size
        # Calculate the width and height of each section
        section_width = width // 4
        section_height = height // 2
        # Iterate over the sections
        for i in range(4):  # Number of sections horizontally
            for j in range(2):  # Number of sections vertically
                # Calculate the coordinates for cropping each section
                left = i * section_width
                upper = j * section_height
                right = (i + 1) * section_width
                lower = (j + 1) * section_height
                # Crop the section
                section = img.crop((left, upper, right, lower))
                # Save the section as a new image
                section.save(output_image_folder + "overview_" + str(i) + "_" + str(j) + ".jpg")


def remove_leading_spaces(text):
    pattern = r"^\s+"
    return re.sub(pattern, "", text, flags=re.MULTILINE)


def escape_file_path(file_path):
    # Replace backslashes with double backslashes
    escaped_path = re.sub(r'\\', r'\\\\', file_path)
    return escaped_path


# Function to process the PDF file
def process_pdf(pdf_path):
    print(f"Processing PDF: {pdf_path}")

    # Values to collect from PDF
    search_terms = ["Spots", "Wrinkles", "Texture", "Pores", "UV Spots", "Brown Spots", "Red Areas", "Porphyrins"]
    # Extract text from PDF
    extracted_text = extract_text_from_pdf(pdf_path)
    # Remove Leading spaces
    cleaned_text = remove_leading_spaces(extracted_text)
    # Find all occurrences of "Each Search Term (<content>)"
    extracted_values = extract_values(cleaned_text, search_terms)

    # Store all the data in a dictionary called data
    if extracted_values is not None:
        data = {term: value for term, value in zip(search_terms, extracted_values)}
    else:
        raise ValueError("No Values Found from PDF")

    # Add the date to the dictionary
    data["date"] = get_report_date(cleaned_text)
    data["name"] = get_report_name(cleaned_text)

    return data


# Get the name on the report
def get_report_name(text):
    pattern = r"face\.(.+?)\s+\."
    matches = re.findall(pattern, text, re.MULTILINE)
    if len(matches) > 1:
        raise ValueError(f"Multiple names found. Ambiguous result.")
    elif matches:
        name = matches[0]
        return name
    else:
        raise ValueError(f"No name found.")


# Get's the date from the report
def get_report_date(text):
    pattern = r"session:\s+([\d/]+)\s+"
    matches = re.findall(pattern, text, re.MULTILINE)
    if len(matches) > 1:
        raise ValueError(f"Multiple dates found. Ambiguous result.")
    elif matches:
        date = matches[0]
        return date
    else:
        raise ValueError(f"No date found.")


# Extracts the values following specified search terms within parentheses.
def extract_values(text, search_terms):
    extracted_values = []

    for search_term in search_terms:
        matches = re.findall(rf"^{search_term}\s*\((\d+\.*\d*)\)", text, re.MULTILINE)
        if len(matches) > 1:
            raise ValueError(f"Multiple '{search_term}' data found. Ambiguous result.")
        elif matches:
            extracted_values.append(matches[0])
        else:
            raise ValueError(f"None of '{search_term}' data found.")

    return extracted_values


def extract_text_from_pdf(pdf_path):
    # Open the PDF file in read-binary mode
    with open(pdf_path, 'rb') as file:
        # Create a PDF file reader object
        pdf_reader = PyPDF2.PdfReader(file)

        # Initialize an empty string to store the extracted text
        text = ""

        # Iterate through each page of the PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

    return text


def check_and_create_folder(folder_path):
    # Check if the folder exists
    if not os.path.exists(folder_path):
        # If the folder does not exist, create it
        os.makedirs(folder_path)
        print(f"Folder '{folder_path}' created.")
    else:
        print(f"Folder '{folder_path}' already exists.")


# Replace forward slashes & spaces with underscores
def format_filename(folder_path):
    folder_name = folder_path.replace('/', '_')
    folder_name = folder_name.replace(' ', '_')
    return folder_name


def add_header_footer(report, page_num):
    # Set the width of the line
    report.setLineWidth(0.5)
    # Colours are RGB colours / 256
    report.setStrokeColorRGB(0.34375, 0.296875, 0.4375)

    # Draw header line
    x_start, y_start = 20 * mm, 270 * mm
    x_end, y_end = 190 * mm, 270 * mm
    report.line(x_start, y_start, x_end, y_end)

    # Draw footer line
    x_start, y_start = 20 * mm, 20 * mm
    x_end, y_end = 190 * mm, 20 * mm
    report.line(x_start, y_start, x_end, y_end)

    # Add the logo
    report.drawImage(escaped_export + "\\Assets\\SkinElementsLogo.jpg", 20 * mm, 274 * mm, 46 * mm, 12.451 * mm)

    # Add the footer text
    report.setFont('OpenSans-Light', 10.5)
    report.setFillColorRGB(0, 0.539, 0.625)
    report.drawString(20 * mm, 12 * mm, str(page_num) + " |", charSpace=1)
    report.setFillColorRGB(0.34375, 0.296875, 0.4375)
    report.drawString(25 * mm, 12 * mm, " SKIN ELEMENTS LIMITED |", charSpace=1)
    report.setFillColorRGB(0, 0.539, 0.625)
    report.drawString(78 * mm, 12 * mm, " FACE UV ANALYSIS ", charSpace=1)

    return report


def euclidean_distance(color1, color2):
    # Calculate the Euclidean distance between two RGB colors.
    r1, g1, b1 = color1
    r2, g2, b2 = color2
    return math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2)


# --- CONFIG --- #

# Folder containing report files
report_folder_path = r'C:\Users\Owner\Downloads\New folder (21)'
# Report Folder Structure
report_export_folder = r'C:\Users\Owner\Downloads\New folder (22)'
# Image Size
size = (2719, 3508)

# --- PROGRAM --- #

try:

    # --- PREPARE --- #

    # Escape the folder path
    escaped_files = escape_file_path(report_folder_path)
    # Escape the export folder location
    escaped_export = escape_file_path(report_export_folder)

    # Check if there are at least two JPG files and at least one PDF file
    check_files(escaped_files)

    # Collect all information from PDF Report
    pdf_files = glob.glob(os.path.join(escaped_files, '*.pdf'))
    for pdf_file in pdf_files:
        data = process_pdf(pdf_file)

    # Create folder for Customer if not existing
    check_and_create_folder(fr"{escaped_export}\\" + format_filename(data["name"]))
    # Create folder for this customers scan today as the date
    check_and_create_folder(
        fr"{escaped_export}\\" + format_filename(data["name"]) + "\\" + format_filename(data["date"]))

    # Create variables for the export folder and pdf
    export_folder = fr"{escaped_export}\\" + format_filename(data["name"]) + "\\" + format_filename(
        data["date"]) + "\\files\\"
    export_pdf = fr"{escaped_export}\\" + format_filename(data["name"]) + "\\" + format_filename(
        data["date"]) + "\\" + format_filename(data["name"]) + "_Report.pdf"

    # Create the export folder
    check_and_create_folder(export_folder)

    # --- RESIZE IMAGES & MOVE FILES --- #

    i = 1
    substring = "overview"
    jpg_files = glob.glob(os.path.join(escaped_files, '*.jpg'))
    for jpg_file in jpg_files:

        # Dont process any image labelled overview
        if substring not in jpg_file:

            # Create a resized version of the image
            resize_image(jpg_file, export_folder + "Image_" + str(i) + ".jpg", size)
            # Progress Indicator
            print(fr"Resized image {jpg_file}")
            # Move the original file
            # --- UNDONE MOVE FOR DEV --- #
            # shutil.move(jpg_file, export_folder + "Raw_Image_" + str(i) + ".jpg")
            # print(fr"Moved image {jpg_file}")

            i = i + 1

        # This is the overview with 8 images
        else:

            # Split this image into 8 sections
            split_image(jpg_file, export_folder)

    for pdf_file in pdf_files:
        # Move the pdf file to the new destination
        # --- UNDONE MOVE FOR DEV --- #
        # shutil.move(pdf_file, export_folder + "raw_report.pdf")
        # print(fr"Moved pdf {pdf_file}")
        # Remove below line
        i = i

    # --- CREATE PHOTO PAGES --- #

    # Create the page
    report = canvas.Canvas(export_pdf, pagesize=A4)
    # Set width and height variables for use
    w, h = A4
    # Register Fonts
    pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
    pdfmetrics.registerFont(TTFont('Calibri Bold', 'Calibrib.ttf'))
    pdfmetrics.registerFont(TTFont('OpenSans-Bold', escaped_export + "\\Assets\\OpenSans-Bold.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-BoldItalic', escaped_export + "\\Assets\\OpenSans-BoldItalic.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-ExtraBold', escaped_export + "\\Assets\\OpenSans-ExtraBold.ttf"))
    pdfmetrics.registerFont(
        TTFont('OpenSans-ExtraBoldItalic', escaped_export + "\\Assets\\OpenSans-ExtraBoldItalic.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-Italic', escaped_export + "\\Assets\\OpenSans-Italic.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-Light', escaped_export + "\\Assets\\OpenSans-Light.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-LightItalic', escaped_export + "\\Assets\\OpenSans-LightItalic.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-Medium', escaped_export + "\\Assets\\OpenSans-Medium.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-MediumItalic', escaped_export + "\\Assets\\OpenSans-MediumItalic.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-Regular', escaped_export + "\\Assets\\OpenSans-Regular.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-SemiBold', escaped_export + "\\Assets\\OpenSans-SemiBold.ttf"))
    pdfmetrics.registerFont(TTFont('OpenSans-SemiBoldItalic', escaped_export + "\\Assets\\OpenSans-SemiBoldItalic.ttf"))
    # Set some colours
    black = (0, 0, 0)
    dark_grey = (0.1, 0.1, 0.1)
    white = (255, 255, 255)
    skn_blue = (0, 0.539, 0.625)
    skn_purple = (0.34375, 0.296875, 0.4375)

    # Create pages for the images and place in the image on them
    i = 1
    for jpg_file in jpg_files:

        # Dont create pages for overview images
        if substring not in jpg_file:

            # Open the image file
            with Image.open(export_folder + "Image_" + str(i) + ".jpg") as img:
                # Pick a single pixel of colour from the centre middle (the chin rest)
                colour_sample = img.getpixel((1240, 3507))

                # Is the colour closer to white or black
                distance1 = euclidean_distance(colour_sample, black)
                distance2 = euclidean_distance(colour_sample, white)

            # Closest Colour is Black, Therefore this is the UV Image
            if distance1 < distance2:

                # Add the image, width and height are in points converted to mm
                report.drawImage(export_folder + "Image_" + str(i) + ".jpg", 0, 0 * mm, 210 * mm, 297 * mm)

                # Place a rectangle on top of the image
                report.setFillColorRGB(1, 1, 1, alpha=0.5)
                report.rect(0, 0, 45 * mm, 6 * mm, fill=True, stroke=False)

                # Add the UV Spots Data
                report.setFont('OpenSans-Bold', 9)
                report.setFillColorRGB(0, 0, 0)
                report.drawString(8 * mm, 1.5 * mm, "Feature Count: " + data["UV Spots"])


            # Closest Colour is White, Therefore this is a regular image
            else:

                # Add the image, width and height are in points converted to mm
                report.drawImage(export_folder + "Image_" + str(i) + ".jpg", 0, 0 * mm, 210 * mm, 297 * mm)

            # Move to the next page
            report.showPage()

            # Iterate to next image
            i = i + 1

    # --- FIRST REPORT PAGE --- #

    styles = getSampleStyleSheet()
    # Create a custom style based on an existing style
    body_style = ParagraphStyle(
        name='Body-1',
        parent=styles['Normal'],
        fontName='OpenSans-Regular',
        wordWrap='default',
        textColor=colors.Color(*dark_grey),
        fontSize=9,
    )
    styles.add(body_style)

    # Set a page number for reference in footer
    page_num = 1
    # Add the header and footer to this page
    report = add_header_footer(report, page_num)

    # Title Section
    report.setFont('OpenSans-ExtraBold', 32)
    report.setFillColorRGB(*skn_blue)
    report.drawString(20 * mm, 252 * mm, "UV Face Analysis")

    report.setFont('OpenSans-BoldItalic', 15)
    report.setFillColorRGB(*skn_purple)
    report.drawString(20 * mm, 243 * mm, "Skin Elements Limited")

    # Paitent Name & Date Section
    report.setFont('OpenSans-Bold', 13.75)
    report.setFillColorRGB(*black)
    report.drawString(20 * mm, 228 * mm, "Paitent Name:")
    report.drawString(100 * mm, 228 * mm, data['name'])

    report.setFont('OpenSans-Regular', 11)
    report.setFillColorRGB(*black)
    report.drawString(20 * mm, 222 * mm, "Analysis Date:")
    report.drawString(100 * mm, 222 * mm, data['date'])

    # Add Photos
    # report.drawImage(export_folder + "overview_0_0.jpg", 30*mm, 169*mm, 34.106*mm, 43.994*mm)
    # report.drawImage(export_folder + "overview_1_0.jpg", 30*mm, 121*mm, 34.106*mm, 43.994*mm)
    # report.drawImage(export_folder + "overview_2_0.jpg", 30*mm, 73*mm, 34.106*mm, 43.994*mm)
    # report.drawImage(export_folder + "overview_3_0.jpg", 30*mm, 25*mm, 34.106*mm, 43.994*mm)

    # Add Photos
    report.drawImage(export_folder + "overview_0_0.jpg", 20 * mm, 129 * mm, 63.677 * mm, 82 * mm)
    report.drawImage(export_folder + "overview_1_0.jpg", 20 * mm, 37 * mm, 63.677 * mm, 82 * mm)

    # Spot Analysis

    spot_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    spot_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.setFont('OpenSans-ExtraBold', 17)
    report.setFillColorRGB(*skn_blue)
    report.drawString(90 * mm, 201 * mm, "Spot Analysis")

    report.setFont('OpenSans-Bold', 12)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 190 * mm, "Primary Spots Identified:")
    report.drawString(170 * mm, 190 * mm, data['Spots'])

    spot_analysis_p1 = Paragraph(spot_analysis_text, styles["Body-1"])
    spot_analysis_p1.wrapOn(report, 90 * mm, 30 * mm)
    spot_analysis_p1.drawOn(report, 90 * mm, 170 * mm)

    report.setFont('OpenSans-Italic', 10)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 163 * mm, "Subheading Goes Here:")

    spot_analysis_p2 = Paragraph(spot_analysis_text_1, styles["Body-1"])
    spot_analysis_p2.wrapOn(report, 90 * mm, 30 * mm)
    spot_analysis_p2.drawOn(report, 90 * mm, 147 * mm)

    # Wrinkle Analysis
    wrinkle_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    wrinkle_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.setFont('OpenSans-ExtraBold', 17)
    report.setFillColorRGB(*skn_blue)
    report.drawString(90 * mm, 105 * mm, "Wrinkle Analysis")

    report.setFont('OpenSans-Bold', 12)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 95 * mm, "Wrinkles Identified:")
    report.drawString(170 * mm, 95 * mm, data['Wrinkles'])

    wrinkle_analysis_p1 = Paragraph(spot_analysis_text, styles["Body-1"])
    wrinkle_analysis_p1.wrapOn(report, 90 * mm, 30 * mm)
    wrinkle_analysis_p1.drawOn(report, 90 * mm, 75 * mm)

    report.setFont('OpenSans-Italic', 10)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 68 * mm, "Subheading Goes Here:")

    wrinkle_analysis_p2 = Paragraph(spot_analysis_text_1, styles["Body-1"])
    wrinkle_analysis_p2.wrapOn(report, 90 * mm, 30 * mm)
    wrinkle_analysis_p2.drawOn(report, 90 * mm, 52 * mm)

    # Move to the next page
    report.showPage()

    # --- SECOND REPORT PAGE --- #

    # Set a page number for reference in footer
    page_num = 2
    # Add the header and footer to this page
    report = add_header_footer(report, page_num)

    # Title Section
    report.setFont('OpenSans-ExtraBold', 32)
    report.setFillColorRGB(*skn_blue)
    report.drawString(20 * mm, 252 * mm, "UV Face Analysis")

    report.setFont('OpenSans-BoldItalic', 15)
    report.setFillColorRGB(*skn_purple)
    report.drawString(20 * mm, 243 * mm, "Skin Elements Limited")

    # Paitent Name & Date Section
    report.setFont('OpenSans-Bold', 13.75)
    report.setFillColorRGB(*black)
    report.drawString(20 * mm, 228 * mm, "Paitent Name:")
    report.drawString(100 * mm, 228 * mm, data['name'])

    report.setFont('OpenSans-Regular', 11)
    report.setFillColorRGB(*black)
    report.drawString(20 * mm, 222 * mm, "Analysis Date:")
    report.drawString(100 * mm, 222 * mm, data['date'])

    # Add Photos
    report.drawImage(export_folder + "overview_2_0.jpg", 20 * mm, 129 * mm, 63.677 * mm, 82 * mm)
    report.drawImage(export_folder + "overview_3_0.jpg", 20 * mm, 37 * mm, 63.677 * mm, 82 * mm)

    # Texture Analysis
    texture_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    texture_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.setFont('OpenSans-ExtraBold', 17)
    report.setFillColorRGB(*skn_blue)
    report.drawString(90 * mm, 201 * mm, "Texture Analysis")

    report.setFont('OpenSans-Bold', 12)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 190 * mm, "Texture Abnormalities:")
    report.drawString(170 * mm, 190 * mm, data['Texture'])

    texture_analysis_p1 = Paragraph(spot_analysis_text, styles["Body-1"])
    texture_analysis_p1.wrapOn(report, 90 * mm, 30 * mm)
    texture_analysis_p1.drawOn(report, 90 * mm, 170 * mm)

    report.setFont('OpenSans-Italic', 10)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 163 * mm, "Subheading Goes Here:")

    texture_analysis_p2 = Paragraph(texture_analysis_text_1, styles["Body-1"])
    texture_analysis_p2.wrapOn(report, 90 * mm, 30 * mm)
    texture_analysis_p2.drawOn(report, 90 * mm, 147 * mm)

    # Pore Analysis
    pore_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    pore_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.setFont('OpenSans-ExtraBold', 17)
    report.setFillColorRGB(*skn_blue)
    report.drawString(90 * mm, 105 * mm, "Pore Analysis")

    report.setFont('OpenSans-Bold', 12)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 95 * mm, "Enlarged Pores:")
    report.drawString(170 * mm, 95 * mm, data['Pores'])

    pore_analysis_p1 = Paragraph(pore_analysis_text, styles["Body-1"])
    pore_analysis_p1.wrapOn(report, 90 * mm, 30 * mm)
    pore_analysis_p1.drawOn(report, 90 * mm, 75 * mm)

    report.setFont('OpenSans-Italic', 10)
    report.setFillColorRGB(*dark_grey)
    report.drawString(90 * mm, 68 * mm, "Subheading Goes Here:")

    pore_analysis_p2 = Paragraph(pore_analysis_text_1, styles["Body-1"])
    pore_analysis_p2.wrapOn(report, 90 * mm, 30 * mm)
    pore_analysis_p2.drawOn(report, 90 * mm, 52 * mm)

    # Move to the next page
    report.showPage()

    # --- THIRD REPORT PAGE --- #

    # Set a page number for reference in footer
    page_num = 3
    # Add the header and footer to this page
    report = add_header_footer(report, page_num)

    # Title Section
    report.setFont('OpenSans-ExtraBold', 32)
    report.setFillColorRGB(*skn_blue)
    report.drawString(20 * mm, 252 * mm, "Important UV Information")

    report.setFont('OpenSans-BoldItalic', 15)
    report.setFillColorRGB(*skn_purple)
    report.drawString(20 * mm, 243 * mm, "Skin Elements Limited")

    # Paitent Name & Date Section
    report.setFont('OpenSans-Bold', 13.75)
    report.setFillColorRGB(*black)
    report.drawString(20 * mm, 228 * mm, "Heading Goes Here")

    # Paragraph 1
    info_para1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive. We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive. We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive. We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    pore_analysis_p1 = Paragraph(info_para1, styles["Body-1"])
    pore_analysis_p1.wrapOn(report, 90 * mm, 30 * mm)
    pore_analysis_p1.drawOn(report, 90 * mm, 75 * mm)

    # Save the report (Overrides)
    report.save()

except ValueError as e:
    print(f"Error: {e}")
