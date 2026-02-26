"""Flask web app for St. Luke worship script to PowerPoint converter.

Upload a .docx worship script, download a .pptx presentation.
"""

import os
import uuid
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

from config import UPLOAD_DIR, OUTPUT_DIR, TEMPLATE_PATH
from docx_parser import parse_worship_script
from slide_generator import SlideGenerator

app = Flask(__name__)
app.secret_key = os.urandom(24)

ALLOWED_EXTENSIONS = {'docx'}

# Ensure directories exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render the upload form."""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    """Accept a .docx upload and return a .pptx download."""
    if 'file' not in request.files:
        flash('No file selected.')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('No file selected.')
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash('Please upload a .docx file.')
        return redirect(url_for('index'))

    # Save uploaded file
    unique_id = str(uuid.uuid4())[:8]
    original_name = secure_filename(file.filename)
    upload_path = os.path.join(UPLOAD_DIR, f'{unique_id}_{original_name}')
    file.save(upload_path)

    try:
        # Parse the worship script
        sections = parse_worship_script(upload_path)

        if not sections:
            flash('Could not parse any sections from the document. '
                  'Please check the format.')
            return redirect(url_for('index'))

        # Generate the presentation
        gen = SlideGenerator()
        gen.generate(sections)

        # Save output
        output_name = original_name.rsplit('.', 1)[0] + '.pptx'
        output_path = os.path.join(OUTPUT_DIR, f'{unique_id}_{output_name}')
        gen.save(output_path)

        slide_count = len(gen.prs.slides)

        # Send the file for download
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        )

    except Exception as e:
        flash(f'Error converting file: {str(e)}')
        return redirect(url_for('index'))

    finally:
        # Clean up uploaded file
        if os.path.exists(upload_path):
            os.remove(upload_path)


if __name__ == '__main__':
    app.run(debug=True, port=8080)
