import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

# Input and template file paths
excel_file = "Book1.xlsx"
template = "template.png"

# Output folder
output_folder = "Generate_certificate"

# Load the Excel file
data = pd.read_excel(excel_file)

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Font settings
bold_font = "arialbd.ttf"
font_name = ImageFont.truetype(bold_font, 90)

# Generate certificates
for index, row in data.iterrows():
        name = row["student_name"]
        certificate = Image.open(template)
        draw = ImageDraw.Draw(certificate)
        if len(name) >= 17 and len(name) < 25:
            name_position = (480, 620)
        elif len(name) >= 14 and len(name) < 16:
            name_position = (550, 620)
        elif len(name) >= 10 and len(name) <14:
            name_position = (620 , 620)
        elif len(name) == "CHHUONG KIMCHHIK":
            name_position == (720, 620)
        elif len(name) >=5 and len(name) <10:
            name_position = (760, 620)
        else:
            name_position = (830, 620)

   
        draw.text(name_position, name, fill="orange", font=font_name)

        # Output path using concatenation
        output_path = output_folder + os.sep + "template_" + name + ".png"

        # Save the certificate
        certificate.save(output_path)

        print("Certificate generated for " + name + " and saved to " + output_path)

print("All certificates have been generated!")