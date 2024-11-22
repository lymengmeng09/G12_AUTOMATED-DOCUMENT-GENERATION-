import openpyxl  # To read Excel files
from docxtpl import DocxTemplate  # To fill Word templates with data
from docx2pdf import convert  # To convert Word documents to PDFs
import os  # To handle file paths and directories


def load_excel_data(excel_file):
    # """Load data from the Excel file."""
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    return list(sheet.values)  # Convert all rows into a list of tuples

def setup_output_directories():
    # """Set up directories for Word and PDF files."""
    word_dir = os.path.join("word_files")
    pdf_dir = os.path.join("pdf_files")
    os.makedirs(word_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    return word_dir, pdf_dir

def generate_word_certificate(template, data, word_dir):
    # """Generate a Word certificate for a given data row."""
    safe_name = str(data["name_e"]).strip().replace("/", "_").replace("\\", "_")
    word_path = os.path.join(word_dir, f"{safe_name}.docx")
    template.render(data)
    template.save(word_path)
    print(f"Created Word document: {word_path}")
    return word_path


def convert_to_pdf(word_path, pdf_dir):
    # """Convert a Word document to PDF."""
    safe_name = os.path.splitext(os.path.basename(word_path))[0]
    pdf_path = os.path.join(pdf_dir, f"{safe_name}.pdf")
    convert(word_path, pdf_path)
    print(f"Created PDF document: {pdf_path}")


def generate_certificates(excel_file, template_file):
    # """Main function to generate certificates."""
    data = load_excel_data(excel_file)
    word_dir, pdf_dir = setup_output_directories()

    template = DocxTemplate(template_file)

    for row in data[1:]:  # Skip the header row
        if not row[1]:  # Skip rows with missing names
            print(f"Skipping row with missing name: {row}")
            continue

        template_data = {
            "name_kh": row[0], "name_e": row[1], "g1": row[2], "g2": row[3],
            "id_kh": row[4], "id_e": row[5], "dob_kh": row[6], "dob_e": row[7],
            "pro_kh": row[8], "pro_e": row[9], "ed_kh": row[10], "ed_e": row[11]
        }

        # Generate Word certificate and convert to PDF
        word_path = generate_word_certificate(template, template_data, word_dir)
        convert_to_pdf(word_path, pdf_dir)

    print("All certificates have been successfully generated!")
# Usage
generate_certificates("Book1.xlsx", "Certificate.docx")
