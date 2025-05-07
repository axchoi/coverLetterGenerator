from docx import Document
from docx2pdf import convert
import os

def customize_and_export(company_name, job_title, output_folder):
    template_path = "cover_letter.docx"
    doc = Document(template_path)

    for para in doc.paragraphs:
        para.text = para.text.replace("[Company Name]", company_name)
        para.text = para.text.replace("[Job Title]", job_title)

    output_docx = os.path.join(output_folder, f"{company_name}_{job_title}_cover_letter.docx")
    doc.save(output_docx)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(base_dir, "Cover_Letters")
    # Convert to PDF
    convert(output_docx)

    print(f"\n✅ Cover letter customized and saved as PDF in:\n {output_folder}")

if __name__ == "__main__":
    print("🔧 Cover Letter Generator\n")
    company_name = input("Enter the company name: ").strip()
    job_title = input("Enter the job title (e.g., Software Engineer): ").strip()
    # Set fixed output folder
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(base_dir, "Cover_Letters")
    # Ensure the folder exists
    os.makedirs(output_folder, exist_ok=True)

    customize_and_export(company_name, job_title, output_folder)