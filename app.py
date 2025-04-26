import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import subprocess
import sys
import shutil

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'  # Required for flash messages
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Temporary folder for uploaded files
app.config['GENERATED_DOCS_FOLDER'] = 'generated_docs/'  # Folder for the Word template
app.config['IMAGE_FOLDER'] = 'path/'  # Folder for the uploaded images
app.config['OUTPUT_FILE'] = 'Final_output.docx'  # Output file path

# Allowed file extensions
ALLOWED_EXCEL_EXTENSIONS = {'xlsx'}
ALLOWED_DOC_EXTENSIONS = {'docx'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg'}

# Ensure upload, generated_docs, and path folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_DOCS_FOLDER'], exist_ok=True)
os.makedirs(app.config['IMAGE_FOLDER'], exist_ok=True)

def allowed_file(filename, allowed_extensions):
    """Check if the file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/')
def index():
    """Render the main page with the upload form."""
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_document():
    """Handle file uploads and document generation."""
    # Check if all required files are part of the request
    if 'excel_file' not in request.files or 'template_file' not in request.files:
        flash('Please upload the Excel sheet and the Word template!')
        return redirect(url_for('index'))

    excel_file = request.files['excel_file']
    template_file = request.files['template_file']

    # Validate Excel and template file uploads
    if excel_file.filename == '' or template_file.filename == '':
        flash('Please select both the Excel sheet and the Word template!')
        return redirect(url_for('index'))

    if not (allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS) and 
            allowed_file(template_file.filename, ALLOWED_DOC_EXTENSIONS)):
        flash('Invalid file types! Excel sheet must be .xlsx, and the template must be .docx.')
        return redirect(url_for('index'))

    # Save the Excel and template files
    excel_filename = secure_filename(excel_file.filename)
    template_filename = secure_filename(template_file.filename)

    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data_ples.xlsx')
    template_path = os.path.join(app.config['GENERATED_DOCS_FOLDER'], 'document_1.docx')

    excel_file.save(excel_path)
    template_file.save(template_path)

    # Handle the image folder upload
    if 'image_folder' not in request.files:
        flash('Please upload the image folder!')
        return redirect(url_for('index'))

    image_files = request.files.getlist('image_folder')

    # Validate and save image files
    if not image_files or all(file.filename == '' for file in image_files):
        flash('Please select a folder containing images!')
        return redirect(url_for('index'))

    # Clear the existing path/ folder to avoid conflicts
    if os.path.exists(app.config['IMAGE_FOLDER']):
        shutil.rmtree(app.config['IMAGE_FOLDER'])
    os.makedirs(app.config['IMAGE_FOLDER'], exist_ok=True)

    # Save each image file into the path/ folder, preserving the relative path
    for image_file in image_files:
        if image_file and allowed_file(image_file.filename, ALLOWED_IMAGE_EXTENSIONS):
            # Get the relative path of the file within the uploaded folder
            relative_path = image_file.filename
            if relative_path.startswith('path/'):
                relative_path = relative_path[len('path/'):]  # Remove the "path/" prefix
            else:
                relative_path = os.path.basename(relative_path)  # Use only the filename if no path

            # Construct the full path to save the image
            image_save_path = os.path.join(app.config['IMAGE_FOLDER'], relative_path)
            
            # Create any necessary subdirectories
            os.makedirs(os.path.dirname(image_save_path), exist_ok=True)
            
            # Save the image
            image_file.save(image_save_path)
        else:
            flash(f'Skipped invalid file: {image_file.filename}. Only .png, .jpg, and .jpeg files are allowed.')
            continue

    # Run the document generation script with the uploaded file paths
    try:
        result = subprocess.run(
            [sys.executable, 'generate_document.py', excel_path, template_path],
            check=True,
            capture_output=True,
            text=True
        )
        print(f"Document generation output: {result.stdout}")
        if result.stderr:
            print(f"Document generation errors: {result.stderr}")
    except subprocess.CalledProcessError as e:
        flash(f"Error generating document: {e.stderr}")
        return redirect(url_for('index'))

    # Check if the output file was created
    if not os.path.exists(app.config['OUTPUT_FILE']):
        flash('Document generation failed: Output file not created.')
        return redirect(url_for('index'))

    # Provide the generated file for download
    return send_file(
        app.config['OUTPUT_FILE'],
        as_attachment=True,
        download_name='Final_output.docx'
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)