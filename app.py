from flask import Flask, request, render_template
import os
import re
from PyPDF2 import PdfMerger
from datetime import datetime
import platform

if platform.system() == "Windows":
    import pythoncom
    import comtypes.client


app = Flask(__name__)

def correct_path(path):
    path = path.strip("'").strip('"')  # Remove unnecessary quotes
    path = re.sub(r'[\\\/]+', r'\\', path)  # Normalize slashes
    path = os.path.expandvars(path)  # Expand environment variables
    return os.path.abspath(path)  # Convert to absolute path

def get_ppt_files_from_directory(input_dir):
    if not os.path.isdir(input_dir):
        return []

    ppt_files = [os.path.join(input_dir, file) for file in os.listdir(input_dir) 
                 if file.endswith('.pptx') or file.endswith('.ppt')]
    return ppt_files

def convert_ppt_to_pdf(ppt_file, pdf_file):
    
    if platform.system() != "Windows":
        print("PowerPoint-to-PDF conversion is only supported on Windows.")
        return
    
    if not os.path.exists(ppt_file):
        print(f"File not found: {ppt_file}")
        return

    try:
        pythoncom.CoInitialize()
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        ppt = powerpoint.Presentations.Open(ppt_file)
        ppt.SaveAs(pdf_file, 32)
        ppt.Close()
        print(f"Converted: {ppt_file}")
    except Exception as e:
        print(f"Error converting {ppt_file}: {e}")
    finally:
        if 'powerpoint' in locals():
            powerpoint.Quit()
        pythoncom.CoUninitialize()

def merge_pdfs(pdf_files, output_path):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()
    print(f"Merged PDF saved as: {output_path}")

@app.route('/', methods=['GET'])
def home():
    return render_template('index.html')  # Show the HTML form

@app.route('/process_folder', methods=['POST'])
def process_folder():
    folder_path = request.form['folder_path']
    folder_path = correct_path(folder_path)  # Correct the provided path

    ppt_files = get_ppt_files_from_directory(folder_path)
    if not ppt_files:
        return f"No PPT files found in the specified folder: {folder_path}"

    output_dir = os.path.join(folder_path, 'PDFs')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert PPT to PDF
    for ppt_file in ppt_files:
        pdf_file = os.path.join(output_dir, os.path.basename(ppt_file).replace('.pptx', '.pdf'))
        convert_ppt_to_pdf(ppt_file, pdf_file)

    # Ask user if they want to merge PDFs
    return render_template('merge.html', folder_path=folder_path)

@app.route('/merge_pdfs', methods=['POST'])
def merge_pdfs_view():
    folder_path = request.form['folder_path']
    folder_path = correct_path(folder_path)  # Correct the provided path

    # Target the PDFs subfolder
    pdf_folder = os.path.join(folder_path, 'PDFs')
    
    # Validate if the PDF folder exists
    if not os.path.exists(pdf_folder):
        return f"No PDF folder found at: {pdf_folder}"

    # Get all PDF files in the 'PDFs' folder
    pdf_files = [os.path.join(pdf_folder, f) for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    if pdf_files:
        # Generate the output merged PDF file path
        output_pdf_name = f"merged_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        output_pdf_path = os.path.join(folder_path, output_pdf_name)
        
        # Merge the PDF files
        merge_pdfs(pdf_files, output_pdf_path)
        return f"Merged PDF saved as: {output_pdf_path}"
    else:
        return "No PDF files found to merge in the 'PDFs' folder."


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

