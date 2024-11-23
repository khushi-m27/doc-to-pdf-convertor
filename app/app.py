from flask import Flask, request, render_template, send_file
from docx import Document
import pypandoc
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file uploaded", 400
    file = request.files['file']
    if file.filename.endswith('.docx'):
        # Read the Word document
        doc = Document(file)
        
        # Extract metadata
        metadata = doc.core_properties
        metadata_info = {
            "author": metadata.author,
            "created": metadata.created,
            "title": metadata.title
        }
        
        # Convert to PDF while preserving formatting
        input_file = BytesIO(file.read())  # Read file into memory
        input_file.seek(0)
        output_pdf = BytesIO()  # Memory buffer for PDF

        try:
            # Convert DOCX to PDF using Pandoc
            pypandoc.convert_file(file.filename, 'pdf', format='docx', outputfile=output_pdf)
        except Exception as e:
            return f"Error during conversion: {str(e)}", 500

        output_pdf.seek(0)  # Reset buffer position to the start

        # Return metadata and allow downloading the file
        return {
            "metadata": metadata_info,
        }, send_file(output_pdf, as_attachment=True, download_name=file.filename.replace('.docx', '.pdf'))

    else:
        return "Invalid file type. Please upload a .docx file.", 400

if __name__ == '__main__':
    app.run(debug=True)
