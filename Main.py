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

                report.add_img(0, 297, 210, 297, export_folder + "Image_" + str(i) + ".jpg")
                report.add_rect(0, 8, 47, 8, report.colours["white"], 0.5, True, False)

                # Add the UV Spots Data
                report.add_text(8.5, 8, "Feature Count: " + data["UV Spots"], "SKN-Caption")

            # Closest Colour is White, Therefore this is a regular image
            else:

                # Add the image, width and height are in points converted to mm
                report.add_img(0, 297, 210, 297, export_folder + "Image_" + str(i) + ".jpg")
                # report.drawImage(export_folder + "Image_" + str(i) + ".jpg", 0, 0 * mm, 210 * mm, 297 * mm)

            # Move to the next page
            report.add_page()

            # Iterate to next image
            i = i + 1

    # --- FIRST REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    # Title Section
    report.add_text(20, 252, "UV Photo-Aging", "SKN-Title")
    report.add_text(20, 243, "Visible Premature Aging - Primary", "SKN-Sub-Title")

    page_summary_text = "UV damage is not entirely invisible. While our UV photo gives you a look under the surface, damage accumulates and slowly effects the upper layers of the skin becoming visible over time. This is known as photo-aging."
    report.add_para(20, 234, 170, 10, page_summary_text, "SKN-Body")

    report.add_text(20, 220, "Patient Name:", "SKN-Heading")
    report.add_text(100, 220, data['name'], "SKN-Heading")

    report.add_text(20, 214, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(100, 214, data['date'], "SKN-Sub-Heading")

    # Add Photos
    report.add_img(20, 201, 63.677, 82, export_folder + "overview_0_0.jpg")
    report.add_img(20, 109, 63.677, 82, export_folder + "overview_1_0.jpg")

    # Spot Analysis
    spot_analysis_text = "These are also known as brown spots, liver spots, or hyperpigmentation."
    spot_analysis_text_1 = "Often people think that freckles for example have been on their skin all their lives. However, if you examine baby photos you will find they haven't. These are caused by cumulative UV damage or particularly bad exposure in once instance. Inflammation from various sources often makes these worse as well."

    report.add_text(90, 195, "Spot Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 184, "Age Spots Identified:", "SKN-Heading")
    report.add_text(170, 184, data['Spots'], "SKN-Heading")
    report.add_para(90, 173, 90, 30, spot_analysis_text, "SKN-Body")
    report.add_text(90, 161, "Important Information", "SKN-Sub-Heading")
    report.add_para(90, 154, 90, 30, spot_analysis_text_1, "SKN-Body")

    # Wrinkle Analysis
    wrinkle_analysis_text = "Wrinkles are the result of the breakdown of collagen and elastin proteins in your skin."
    wrinkle_analysis_text_1 = "Collagen and elastin are what make your skin firm, supple, and resilient. UV Radiation directly damages and destroys collagen and elastin leading to wrinkles developing over years of exposure.  Virtually all wrinkles will be found in skin that is regularly exposed to sunlight as this is the primary cause of their development. "

    report.add_text(90, 101, "Wrinkle Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 91, "Wrinkles Identified:", "SKN-Heading")
    report.add_text(170, 91, data['Wrinkles'], "SKN-Heading")
    report.add_para(90, 80, 90, 30, wrinkle_analysis_text, "SKN-Body")
    report.add_text(90, 68, "Important Information:", "SKN-Sub-Heading")
    report.add_para(90, 61, 90, 30, wrinkle_analysis_text_1, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- SECOND REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    report.add_text(20, 252, "UV Photo-Aging", "SKN-Title")
    report.add_text(20, 243, "Visible Premature Aging - Secondary", "SKN-Sub-Title")

    page_summary_text = "UV damage is not entirely invisible. While our UV photo gives you a look under the surface, damage accumulates and slowly effects the upper layers of the skin becoming visible over time. This is known as photo-aging."
    report.add_para(20, 234, 170, 10, page_summary_text, "SKN-Body")

    report.add_text(20, 220, "Patient Name:", "SKN-Heading")
    report.add_text(100, 220, data['name'], "SKN-Heading")

    report.add_text(20, 214, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(100, 214, data['date'], "SKN-Sub-Heading")

    # Add Photos
    report.add_img(20, 201, 63.677, 82, export_folder + "overview_2_0.jpg")
    report.add_img(20, 109, 63.677, 82, export_folder + "overview_3_0.jpg")

    # Texture Analysis
    texture_analysis_text = "Texture abnormalities are indications of undesirable patchy skin with a rougher texture."
    texture_analysis_text_1 = "Texture abnormalities usually have the same causes as wrinkles, and can be the precursor for their development. It often also has causes related to moisture, oil loss and inflammation, all common effects of UV exposure. In short skin texture tends to improve as the overall health of your skin improves and any form of damage tends to show in the texture. "

    report.add_text(90, 196, "Texture Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 185, "Texture Abnormalities:", "SKN-Heading")
    report.add_text(170, 185, data['Texture'], "SKN-Heading")

    report.add_para(90, 174, 90, 30, texture_analysis_text, "SKN-Body")
    report.add_text(90, 162, "Important Information:", "SKN-Sub-Heading")
    report.add_para(90, 155, 90, 30, texture_analysis_text_1, "SKN-Body")

    # Pore Analysis
    pore_analysis_text = "Larger than normal pores individually identified in our analysis."
    pore_analysis_text_1 = "Enlarged skin pores are caused by a couple of main causes. Sun damage is the first, UV radiation affecting the collagen often results in pores becoming larger and more open. The other common cause is increased oil production, which results in oily skin and enlarged pores. It should also be noted that increased oil production is another symptom of UV damage."

    report.add_text(90, 103, "Pore Analysis", "SKN-Coloured-Heading")

    report.add_text(90, 93, "Enlarged Pores:", "SKN-Heading")
    report.add_text(170, 93, data['Pores'], "SKN-Heading")

    report.add_para(90, 82, 90, 30, pore_analysis_text, "SKN-Body")
    report.add_text(90, 70, "Important Information:", "SKN-Sub-Heading")
    report.add_para(90, 63, 90, 30, pore_analysis_text_1, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- THIRD REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    report.add_text(20, 252, "UV Radiation", "SKN-Title")
    report.add_text(20, 243, "Preventing UV Exposure", "SKN-Sub-Title")

    info_para1 = "Photoaging is enemy number one for every skin cell on your body. UV damage is the primary cause of looking older year after year. The good news is that you don’t just have to sit there and take it. "
    info_para2 = "Your body adapts and responds. With proper protection from radiation, restoration will occur even without assistance. However, it won’t be fast for damage accumulated over many years."
    report.add_para(20, 230, 135, 20, info_para1, "SKN-Body")
    report.add_para(20, 215, 135, 20, info_para2, "SKN-Body")
    report.add_text(20, 200, "The most important thing is that you protect your skin from exposure.", "SKN-Body")

    report.add_img(160, 233, 30, 30, escaped_export + "\\Assets\\LearnMoreUV.png")
    report.add_text(161, 205, "Learn More About", "SKN-Body")
    report.add_text(165, 201, "UV Radiation", "SKN-Body")

    report.add_img(20, 190, 175.37, 87.871, escaped_export + "\\Assets\\UVA-Radiation-Vs-UVB-Radiation.jpg")

    report.add_text(20, 95, "The Dangers of UV-A", "SKN-Coloured-Heading")

    info_para3 = "Unless you are welding or taking a spacewalk, there are only two kinds of UV radiation to pay attention to, UV-A and UV-B."
    info_para4 = "UV-B is what causes your skin to burn from overexposure, this kind of radiation is a shorter wavelength. UV-A is what causes your skin to age, this radiation has a longer wavelength and penetrates deeply into the skin."
    info_para5 = "The problem comes when people don’t understand the difference. UV-B radiation, with it’s shorter wavelength, is fairly easy to stop. A hat, window or even just shade will often suffice. However UV-A radiation is where the real danger lies. "
    report.add_para(20, 85, 83, 20, info_para3, "SKN-Body")
    report.add_para(20, 72, 83, 20, info_para4, "SKN-Body")
    report.add_para(20, 47, 83, 20, info_para5, "SKN-Body")

    info_para6 = "UV-A radiation penetrates through many objects, windows, shade, even heavy cloud cover or rain do not stop it. Even with great sun care routines, people are exposed heavily every single day without ever realizing it. This is where the majority of sun damage comes from. "
    info_para7 = "UV-A even reflects off sidewalks, tiles, snow, and water, to just name a few surfaces, often resulting in much higher levels of exposure just by where you are standing. Often you are exposed just standing in a house with a window in the room. Combine this with the fact that 95% of UV radiation from the sun is UV-A and you find that UV-A is by far the biggest UV threat your body faces."
    report.add_para(106, 85, 83, 20, info_para6, "SKN-Body")
    report.add_para(106, 60, 83, 20, info_para7, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- FOURTH REPORT PAGE --- #

    # Add the header and footer to this page
    report.add_header_footer()

    report.add_text(20, 252, "SPF Ratings & Values", "SKN-Title")
    report.add_text(20, 243, "Preventing UV Exposure", "SKN-Sub-Title")

    info_para1 = "Most people choose a sunscreen based on the SPF value on the front. Unfortunately SPF is a poor measure for choosing sunscreen and has two major problems. "
    info_para2 = "The first problem is that SPF does not measure protection from UV-A radiation, which we have already ascertained is the larger threat to your skin on a daily basis. SPF only measures UV-B protection. It’s far more important you choose a sunscreen marked as broad spectrum with well-chosen ingredients, than to have the highest possible SPF. Because while the higher SPF value may offer marginally better UV-B protection than a lower one, it may offer significantly less UV-A protection. Often the UV-A protection suffers in the marketing efforts of the UV-B protection increase."
    info_para3 = "The second problem is that SPF has no influence on reapplication frequency. People often purchase a higher SPF and wait longer before reapplying, however application frequency is not affected at all by the higher SPF. "
    info_para4 = "A final note on SPF 50. SPF 50 is often seen as the pinnacle of SPF protection in Australia. Unfortunately while SPF 50 sounds like almost double the lower rating of SPF 30, the reality is completely different. Would it surprise you to know that SPF 50 is only 1.3% more effective at blocking UV-B than SPF 30? "
    info_para5 = "You can scan the QR code above to learn more about SPF values and how they are calculated, however suffice to say as shown by the infographic below. SPF is a very deceptive measure of sun protection. "
    report.add_para(60, 230, 125, 30, info_para1, "SKN-Body")
    report.add_para(60, 217, 125, 30, info_para2, "SKN-Body")
    report.add_para(60, 184, 125, 30, info_para3, "SKN-Body")
    report.add_para(60, 168, 125, 30, info_para4, "SKN-Body")
    report.add_para(60, 148, 125, 30, info_para5, "SKN-Body")

    report.add_text(20, 230, "Learn More About", "SKN-Body")
    report.add_text(24, 226, "SPF Ratings", "SKN-Body")
    report.add_img(20, 220, 30, 30, escaped_export + "\\Assets\\LearnMoreSPF.png")

    report.add_img(20, 120, 170, 89.746, escaped_export + "\\Assets\\UVB-Radiation-SPF.jpg")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- FIFTH REPORT PAGE --- #
    """
    report.add_text(20, 252, "Recommendations", "SKN-Title")
    report.add_text(20, 243, "Analysis Summary & Recommendations", "SKN-Sub-Title")

    report.add_text(20, 230, "Patient Name:", "SKN-Heading")
    report.add_text(90, 230, data['name'], "SKN-Heading")

    report.add_text(20, 224, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(90, 224, data['date'], "SKN-Sub-Heading")

    report.add_text(20, 215, "Melanin Spots:", "SKN-Heading")
    report.add_text(90, 215, data['UV Spots'], "SKN-Heading")

    report.add_text(20, 208, "Age Spots:", "SKN-Sub-Heading")
    report.add_text(90, 208, data['Spots'], "SKN-Sub-Heading")
    report.add_text(20, 203, "Wrinkles:", "SKN-Sub-Heading")
    report.add_text(90, 203, data['Wrinkles'], "SKN-Sub-Heading")
    report.add_text(110, 208, "Texture Abnormalities:", "SKN-Sub-Heading")
    report.add_text(175, 208, data['Texture'], "SKN-Sub-Heading")
    report.add_text(110, 203, "Enlarged Pores", "SKN-Sub-Heading")
    report.add_text(175, 203, data['Pores'], "SKN-Sub-Heading")

    report.add_text(20, 185, "Ingredients:", "SKN-Coloured-Heading")
    report.add_text(20, 177, "Ingredients We Recommend", "SKN-Heading")
    report.add_text(20, 170, "Zinc Oxide (Micronised)", "SKN-Sub-Heading")
    report.add_text(20, 165, "Cucumber Extract", "SKN-Sub-Heading")
    report.add_text(20, 160, "Aloe Vera", "SKN-Sub-Heading")
    report.add_text(20, 155, "Argan Oil", "SKN-Sub-Heading")
    report.add_text(20, 150, "Jojoba Oil", "SKN-Sub-Heading")
    report.add_text(20, 145, "Papaya", "SKN-Sub-Heading")
    report.add_text(20, 140, "Geranium Oil", "SKN-Sub-Heading")
    report.add_text(20, 135, "Tea Tree Extract", "SKN-Sub-Heading")
    report.add_text(20, 130, "Rosehip Oil", "SKN-Sub-Heading")
    report.add_text(20, 125, "Safflower Oil", "SKN-Sub-Heading")
    report.add_text(20, 120, "Raspberry", "SKN-Sub-Heading")
    report.add_text(20, 115, "Lavender", "SKN-Sub-Heading")

    report.add_text(110, 177, "We Suggest Avoiding", "SKN-Heading")
    report.add_text(110, 170, "Zinc Oxide (Nano)", "SKN-Sub-Heading")
    report.add_text(110, 165, "Titanium Dioxide", "SKN-Sub-Heading")
    report.add_text(110, 160, "Octinoxate (Octyl-Methoxycinnamate)", "SKN-Sub-Heading")
    report.add_text(110, 155, "Oxybenzone (Benzophenone-3)", "SKN-Sub-Heading")
    report.add_text(110, 150, "Avobenzone (AVO)", "SKN-Sub-Heading")
    report.add_text(110, 145, "PABA (p-aminobenzoic acid)", "SKN-Sub-Heading")
    report.add_text(110, 140, "Cinoxate", "SKN-Sub-Heading")
    report.add_text(110, 135, "Padimate-O", "SKN-Sub-Heading")
    report.add_text(110, 130, "Methyl sinapate", "SKN-Sub-Heading")
    report.add_text(110, 125, "Dibenzoylmethane and Parsol 1789", "SKN-Sub-Heading")
    report.add_text(110, 120, "2-phenylbenzimidazole", "SKN-Sub-Heading")
    report.add_text(110, 115, "4-Methyl-benzylidene-camphor (4-MBC)", "SKN-Sub-Heading")

    report.add_img(105, 103, 108.564, 81.351, escaped_export + "\\Assets\\NaturalSunscreenFaceMoisturising_Finalist.jpg")
    report.add_img(17, 62, 30, 30, escaped_export + "\\Assets\\FACE.png")

    report.add_text(20, 95, "Sunscreen Recommended:", "SKN-Coloured-Heading")
    report.add_text(20, 87, "Soléo Organics Face Moisturising", "SKN-Heading")

    info_para1 = "Based on your scan we recommend this sunscreen for your daily use. This recommendation takes identified UV damage into consideration, and any physical signs of aging identified on the skin."
    report.add_para(20, 80, 100, 30, info_para1, "SKN-Body")

    report.add_text(50, 60, "Features:", "SKN-Heading")
    report.add_text(50, 55, "SPF 30 Protection", "SKN-Sub-Heading")
    report.add_text(50, 50, "Broad Spectrum", "SKN-Sub-Heading")
    report.add_text(50, 45, "Lightly Moisturising", "SKN-Sub-Heading")
    report.add_text(50, 40, "Non-Comedogenic", "SKN-Sub-Heading")

    report.add_text(90, 55, "Anti-Aging", "SKN-Sub-Heading")
    report.add_text(90, 50, "Matte (Non-Shiny)", "SKN-Sub-Heading")
    report.add_text(90, 45, "Rubs in Clear", "SKN-Sub-Heading")

    # Add the header and footer to this page
    report.add_header_footer()
    """

    report.add_header_footer()

    report.add_text(20, 252, "Recommendations", "SKN-Title")
    report.add_text(20, 243, "Analysis Summary & Recommendations", "SKN-Sub-Title")

    report.add_text(20, 230, "Patient Name:", "SKN-Heading")
    report.add_text(90, 230, data['name'], "SKN-Heading")

    report.add_text(20, 224, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(90, 224, data['date'], "SKN-Sub-Heading")

    report.add_text(20, 215, "Melanin Spots:", "SKN-Heading")
    report.add_text(90, 215, data['UV Spots'], "SKN-Heading")

    report.add_text(20, 208, "Age Spots:", "SKN-Sub-Heading")
    report.add_text(90, 208, data['Spots'], "SKN-Sub-Heading")
    report.add_text(20, 203, "Wrinkles:", "SKN-Sub-Heading")
    report.add_text(90, 203, data['Wrinkles'], "SKN-Sub-Heading")
    report.add_text(110, 208, "Texture Abnormalities:", "SKN-Sub-Heading")
    report.add_text(175, 208, data['Texture'], "SKN-Sub-Heading")
    report.add_text(110, 203, "Enlarged Pores", "SKN-Sub-Heading")
    report.add_text(175, 203, data['Pores'], "SKN-Sub-Heading")

    report.add_text(20, 188, "Sunscreen Ingredient Suggestions", "SKN-Coloured-Heading")

    report.add_text(20, 178, "UV Protection:", "SKN-Heading")
    report.add_text(20, 170, "Zinc Oxide (Micronised):", "SKN-Sub-Heading")
    info_para1 = "Broad spectrum UV protection from UV-A I & UVA II. Highly stable in UV light. Anti-oxidant with anti-inflammatory properties."
    report.add_para(20, 165, 80, 30, info_para1, "SKN-Body")
    report.add_text(20, 150, "Geranium Oil:", "SKN-Sub-Heading")
    info_para1 = "Very high in antioxidants. Helps the ward off free radicals. Assist the body in preventing UV damage. Astringent properties can help constrict pores."
    report.add_para(20, 145, 80, 30, info_para1, "SKN-Body")
    report.add_text(20, 130, "Tea Tree Extract:", "SKN-Sub-Heading")
    info_para1 = "Contains catechins and flavonoids which act as sunscreen. Good source of Vitamin E. High in Antioxidants."
    report.add_para(20, 125, 80, 30, info_para1, "SKN-Body")

    report.add_text(110, 178, "Skin Conditioning:", "SKN-Heading")
    report.add_text(110, 170, "Raspberry:", "SKN-Sub-Heading")
    info_para1 = "Highly soothing. HIgh levels of Vitamins A and E. Helps create a lipid barrier to protect the skin from losing moisture."
    report.add_para(110, 165, 80, 30, info_para1, "SKN-Body")
    report.add_text(110, 150, "Aloe Vera:", "SKN-Sub-Heading")
    info_para1 = "Highly soothing and mositurising. Helps counteract UV damage. High in vitamins A, E, and C. Astringent properties can help constrict pores."
    report.add_para(110, 145, 80, 30, info_para1, "SKN-Body")
    report.add_text(110, 130, "Cucumber Extract:", "SKN-Sub-Heading")
    info_para1 = "Highly soothing skin care ingredient. Provides instant moisturisation and strengthens the skins natural mositure barrier."
    report.add_para(110, 125, 80, 30, info_para1, "SKN-Body")

    report.add_text(20, 105, "Photo-Aging Reduction:", "SKN-Heading")

    report.add_text(20, 97, "Jojoba Oil:", "SKN-Sub-Heading")
    info_para1 = "High in anti-oxidants. Helps fight off free radicals. High in vitamin E, and promotes skin cell regeneration."
    report.add_para(20, 92, 80, 30, info_para1, "SKN-Body")
    report.add_text(20, 77, "Argan Oil:", "SKN-Sub-Heading")
    info_para1 = "Gently exfoliator. Assists in the restoration of sun damaged skin. Helps reduce fine lines. High in anti-oxidants."
    report.add_para(20, 72, 80, 30, info_para1, "SKN-Body")
    report.add_text(20, 56, "Rosehip Oil:", "SKN-Sub-Heading")
    info_para1 = "Often used as a non-toxic alternative to retinol. Helps restore elasticity and promotes cell regeneration."
    report.add_para(20, 50, 80, 30, info_para1, "SKN-Body")
    report.add_text(110, 97, "Papaya Extract:", "SKN-Sub-Heading")
    info_para1 = "Promotes skin turnover and renewal. Natural exfoliator. Rejuvenates stressed skin. Great for age spots and poor skin texture."
    report.add_para(110, 92, 80, 30, info_para1, "SKN-Body")
    report.add_text(110, 77, "Safflower Oil:", "SKN-Sub-Heading")
    info_para1 = "High in polyphenols which inhibit enzymes that break down skin proteins like collagen and elastin. High in antioxidants. "
    report.add_para(110, 72, 80, 30, info_para1, "SKN-Body")
    report.add_text(110, 56, "Green Tea Extract:", "SKN-Sub-Heading")
    info_para1 = "Helps prevent the breakdown of collagen and elastin. Reduces irritation and elastin damage. High in antioxidants."
    report.add_para(110, 50, 80, 30, info_para1, "SKN-Body")

    # Move to the next page
    report.add_page()
    report.inc_page_num()

    # --- SIXTH REPORT PAGE --- #

    report.add_header_footer()

    report.add_text(20, 252, "Recommended Sunscreen", "SKN-Title")
    report.add_text(20, 243, "Analysis Summary & Recommendations", "SKN-Sub-Title")

    report.add_text(20, 230, "Patient Name:", "SKN-Heading")
    report.add_text(90, 230, data['name'], "SKN-Heading")

    report.add_text(20, 224, "Analysis Date:", "SKN-Sub-Heading")
    report.add_text(90, 224, data['date'], "SKN-Sub-Heading")

    report.add_img(105, 172, 108.564, 81.351, escaped_export + "\\Assets\\NaturalSunscreenFaceMoisturising_Finalist.jpg")

    # report.add_img(165, 55, 30, 30, escaped_export + "\\Assets\\FACE.png")
    # report.add_text(171, 28, "Learn More", "SKN-Body")

    report.add_text(20, 200, "Sunscreen Recommendation", "SKN-Coloured-Heading")
    report.add_text(20, 190, "We would like to offer the following sunscreen recommendation for use daily.", "SKN-Body")

    report.add_text(20, 175, "Soléo Organics Face Moisturising", "SKN-Heading")
    info_para1 = "This all natural sunscreen, avoids many of the chemicals associated high higher levels of irritation which can be counter productive. "
    report.add_para(20, 165, 100, 30, info_para1, "SKN-Body")
    info_para1 = "Based on a miconised zinc formula it offers strong protection from both UV-A and UV-B radiation, even in the higher UV-A I band. "
    report.add_para(20, 152, 100, 30, info_para1, "SKN-Body")
    info_para1 = "In addition it contains a mixture of other natural extracts to counteract the progress of UV damage on your skin thus far. For example, Rosehip oil to promote cell renewal and turnover, jojoba oil and argan oil to both exfoliate slightly and reduce fine lines and wrinkles that have already formed and safflower oil to reduce the degredation of the skins collagen and elastin. "
    report.add_para(20, 139, 100, 30, info_para1, "SKN-Body")
    info_para1 = "Finally as an all natural formula it helps promote sustainability with a completely biodegradable and reef safe formula and completely recyclable packaging."
    report.add_para(20, 110, 100, 30, info_para1, "SKN-Body")

    report.add_text(20, 90, "Summary:", "SKN-Heading")

    report.add_text(20, 80, "SPF 30 Protection", "SKN-Sub-Heading")
    report.add_text(20, 75, "Broad Spectrum", "SKN-Sub-Heading")
    report.add_text(20, 70, "Lightly Moisturising", "SKN-Sub-Heading")
    report.add_text(20, 65, "Non-Comedogenic", "SKN-Sub-Heading")
    report.add_text(20, 60, "Biodegradable/Reef Safe", "SKN-Sub-Heading")
    report.add_text(20, 55, "Cruelty Free", "SKN-Sub-Heading")

    report.add_text(85, 80, "Anti-Aging", "SKN-Sub-Heading")
    report.add_text(85, 75, "Reduces Fine Lines", "SKN-Sub-Heading")
    report.add_text(85, 70, "Promotes Skin Renewal", "SKN-Sub-Heading")

    report.add_text(145, 80, "Matte (Non Shiny)", "SKN-Sub-Heading")
    report.add_text(145, 75, "Non-Comedogenic", "SKN-Sub-Heading")
    report.add_text(145, 70, "Rubs in Clear", "SKN-Sub-Heading")
    report.add_text(145, 65, "Effective Under Makeup", "SKN-Sub-Heading")

    # Save the report (Overrides)
    report.save_report()

except ValueError as e:
    print(f"Error: {e}")
