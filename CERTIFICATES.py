import openpyxl
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog, messagebox


def select_file(file_type, file_extensions, title):
    """Prompt the user to select a file and return the file path."""
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[(file_type, file_extensions), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showwarning("No File Selected", f"Please select a {file_type.lower()} file.")
    return file_path


def load_excel_file(file_path):
    """Load the Excel file and return the active sheet."""
    try:
        workbook = openpyxl.load_workbook(file_path)
        return workbook.active
    except Exception as e:
        messagebox.showerror("Invalid File", "The selected file is not a valid Excel file.")
        return None


def validate_excel_data(sheet):
    """Validate the structure of the Excel data."""
    data = list(sheet.values)
    if len(data) < 2 or len(data[0]) < 3:
        messagebox.showerror("Invalid File Structure", "The Excel file must contain at least three columns (e.g., student name, teacher name, leader name) with data.")
        return None
    return data


def load_word_template(file_path):
    """Load the Word template and return the DocxTemplate object."""
    try:
        return DocxTemplate(file_path)
    except Exception as e:
        messagebox.showerror("Invalid File", "The selected file is not a valid Word template.")
        return None


def generate_personalized_documents(student_data, word_template):
    """Generate Word documents for each student."""
    for student in student_data[1:]:  # Skip the header row
        try:
            word_template.render({
                'student_name': student[0],
                'teacher_name': student[1],
                'leader_name': student[2],
            })
            output_file = f"{student[0]}.docx"
            word_template.save(output_file)
        except Exception as e:
            messagebox.showerror("Error Generating Document", f"An error occurred while generating the document for {student[0]}: {str(e)}")
            continue


def generate_documents():
    """Main function to generate documents."""
    try:
        # Step 1: Select and validate Excel file
        excel_path = select_file("Excel File", "*.xlsx", "Select the Excel File")
        if not excel_path:
            return

        sheet = load_excel_file(excel_path)
        if not sheet:
            return

        student_data = validate_excel_data(sheet)
        if not student_data:
            return

        # Step 2: Select and validate Word template
        word_template_path = select_file("Word Template", "*.docx", "Select the Word Template File")
        if not word_template_path:
            return

        word_template = load_word_template(word_template_path)
        if not word_template:
            return

        # Step 3: Generate documents
        generate_personalized_documents(student_data, word_template)

        messagebox.showinfo("Success", "Documents generated successfully!")
    except Exception as e:
        messagebox.showerror("Unexpected Error", f"An unexpected error occurred: {str(e)}")


# Tkinter GUI setup
root = tk.Tk()
root.title("Document Generator")
root.geometry("400x200")

# Create and place the button
generate_button = tk.Button(root, text="Generate Documents", command=generate_documents, width=30, height=2, bg="lightblue")
generate_button.pack(pady=50)

# Run the Tkinter event loop
root.mainloop()