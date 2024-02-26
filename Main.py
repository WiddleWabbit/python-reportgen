import os
import re
import glob
import PyPDF2
import shutil
import math

from PIL import Image
from Report import ReportGenerator

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
    report = ReportGenerator(export_pdf, escaped_export + "\\Assets\\")

    # Set some colours
    ###
    black = (0, 0, 0)
    dark_grey = (0.1, 0.1, 0.1)
    white = (255, 255, 255)
    skn_blue = (0, 0.539, 0.625)
    skn_purple = (0.34375, 0.296875, 0.4375)
    ###

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

                report.add_img(0, 0, 210, 297, export_folder + "Image_" + str(i) + ".jpg")
                report.add_rect(0, 0, 45, 6, report.colours["white"], 0.5, True, False)

                # Add the UV Spots Data
                report.add_text(8, 1.5, "Feature Count: " + data["UV Spots"], "SKN-Caption")

            # Closest Colour is White, Therefore this is a regular image
            else:

                # Add the image, width and height are in points converted to mm
                report.add_img(0, 0, 210, 297, export_folder + "Image_" + str(i) + ".jpg")
                # report.drawImage(export_folder + "Image_" + str(i) + ".jpg", 0, 0 * mm, 210 * mm, 297 * mm)

            # Move to the next page
            report.add_page()

            # Iterate to next image
            i = i + 1

    # --- FIRST REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    # Title Section
    report.add_text(20, 252, "UV Face Analysis", "SKN-Title")
    report.add_text(20, 243, "Skin Elements Limited", "SKN-Sub-Title")

    report.add_text(20, 228, "Patient Name:", "SKN-Heading")
    report.add_text(100, 228, data['name'], "SKN-Heading")

    report.add_text(20, 222, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(100, 222, data['date'], "SKN-Sub-Heading")

    # Add Photos
    report.add_img(20, 129, 63.677, 82, export_folder + "overview_0_0.jpg")
    report.add_img(20, 37, 63.677, 82, export_folder + "overview_1_0.jpg")

    # Spot Analysis
    spot_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    spot_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.add_text(90, 201, "Spot Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 190, "Primary Spots Identified:", "SKN-Sub-Heading")
    report.add_text(170, 190, data['Spots'], "SKN-Sub-Heading")
    report.add_para(90, 170, 90, 30, spot_analysis_text, "SKN-Body")
    report.add_text(90, 163, "Subheading Goes Here:", "SKN-Sub-Heading")
    report.add_para(90, 147, 90, 30, spot_analysis_text_1, "SKN-Body")

    # Wrinkle Analysis
    wrinkle_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    wrinkle_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.add_text(90, 105, "Wrinkle Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 95, "Wrinkles Identified:", "SKN-Sub-Heading")
    report.add_text(170, 95, data['Wrinkles'], "SKN-Sub-Heading")
    report.add_para(90, 75, 90, 30, wrinkle_analysis_text, "SKN-Body")
    report.add_text(90, 68, "Subheading Goes Here:", "SKN-Sub-Heading")
    report.add_para(90, 52, 90, 30, wrinkle_analysis_text_1, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- SECOND REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    report.add_text(20, 252, "UV Face Analysis", "SKN-Title")
    report.add_text(20, 243, "Skin Elements Limited", "SKN-Sub-Title")

    report.add_text(20, 228, "Patient Name:", "SKN-Heading")
    report.add_text(100, 228, data['name'], "SKN-Heading")

    report.add_text(20, 222, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(100, 222, data['date'], "SKN-Sub-Heading")

    # Add Photos
    report.add_img(20, 129, 63.677, 82, export_folder + "overview_2_0.jpg")
    report.add_img(20, 37, 63.677, 82, export_folder + "overview_3_0.jpg")

    # Texture Analysis
    texture_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    texture_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.add_text(90, 201, "Texture Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 190, "Texture Abnormalities:", "SKN-Sub-Heading")
    report.add_text(170, 190, data['Texture'], "SKN-Sub-Heading")

    report.add_para(90, 170, 90, 30, texture_analysis_text, "SKN-Body")
    report.add_text(90, 163, "Subheading Goes Here:", "SKN-Sub-Heading")
    report.add_para(90, 147, 90, 30, texture_analysis_text_1, "SKN-Body")

    # Pore Analysis
    pore_analysis_text = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."
    pore_analysis_text_1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to comprehensive."

    report.add_text(90, 105, "Pore Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 95, "Enlarged Pores:", "SKN-Sub-Heading")
    report.add_text(170, 95, data['Pores'], "SKN-Sub-Heading")

    report.add_para(90, 75, 90, 30, pore_analysis_text, "SKN-Body")
    report.add_text(90, 68, "Subheading Goes Here:", "SKN-Sub-Heading")
    report.add_para(90, 52, 90, 30, pore_analysis_text_1, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- THIRD REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    report.add_text(20, 252, "Important UV Information", "SKN-Title")
    report.add_text(20, 243, "Skin Elements Limited", "SKN-Sub-Title")

    # Paitent Name & Date Section
    report.add_text(20, 228, "Heading Goes Here", "SKN-Heading")

    # Paragraph 1
    info_para1 = "We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to omprehensive.We want to empower people with comprehensive knowledge about what they apply to their skin. power people with power people with e want to omprehensive.We want to empower people with comprehensive knowledge about what they apply to their skin. "
    report.add_para(20, 180, 70, 30, info_para1, "SKN-Body")

    report.add_img(100, 175, 87.37, 46.443, escaped_export + "\\Assets\\UVB-Radiation-SPF.jpg")

    # Save the report (Overrides)
    report.save_report()

except ValueError as e:
    print(f"Error: {e}")
