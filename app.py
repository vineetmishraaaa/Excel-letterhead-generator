import os
import subprocess
import tempfile
from io import BytesIO
from pathlib import Path
from threading import Lock

import fitz
from PIL import Image, ImageOps
from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / 'uploads'
OUTPUT_FOLDER = BASE_DIR / 'output'

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ALLOWED_EXCEL = {'.xlsx', '.xls', '.xlsm'}
ALLOWED_LETTERHEAD = {'.pdf'}

LETTERHEAD_PAGE = 0
A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
CONTENT_TOP_CM = 4.15
CONTENT_WIDTH_CM = 19.60
CONTENT_BOTTOM_MARGIN_CM = 0.70
CONTENT_GAP_CM = 0.20

# LibreOffice/Calc should only be driven by one request at a time.
CONVERSION_LOCK = Lock()


def cm_to_pt(cm):
    return cm * 72.0 / 2.54


def get_non_white_bbox(pdf_page):
    pix = pdf_page.get_pixmap(
        matrix=fitz.Matrix(2.0, 2.0),
        colorspace=fitz.csRGB,
        alpha=False,
    )

    image = Image.frombytes(
        'RGB',
        [pix.width, pix.height],
        pix.samples,
    )

    gray = ImageOps.grayscale(image)

    mask = gray.point(
        lambda value: 255 if value < 245 else 0
    )

    bbox = mask.getbbox()

    if bbox is None:
        return pdf_page.rect

    left, top, right, bottom = bbox
    padding_px = 6

    left = max(0, left - padding_px)
    top = max(0, top - padding_px)
    right = min(pix.width, right + padding_px)
    bottom = min(pix.height, bottom + padding_px)

    sx = pdf_page.rect.width / pix.width
    sy = pdf_page.rect.height / pix.height

    return fitz.Rect(
        left * sx,
        top * sy,
        right * sx,
        bottom * sy,
    )


def calculate_destination_rect(page, source_rect):
    page_width = page.rect.width
    page_height = page.rect.height

    max_width = cm_to_pt(CONTENT_WIDTH_CM)
    top = cm_to_pt(CONTENT_TOP_CM + CONTENT_GAP_CM)
    bottom_margin = cm_to_pt(CONTENT_BOTTOM_MARGIN_CM)
    max_height = page_height - top - bottom_margin

    if source_rect.width <= 0 or source_rect.height <= 0:
        raise ValueError('Excel PDF page contains no usable content.')

    width_scale = max_width / source_rect.width
    height_scale = max_height / source_rect.height
    scale = min(width_scale, height_scale)

    dest_width = source_rect.width * scale
    dest_height = source_rect.height * scale

    x = (page_width - dest_width) / 2.0
    y = top

    return fitz.Rect(
        x,
        y,
        x + dest_width,
        y + dest_height,
    )


def export_excel_to_pdf(excel_path, output_pdf):
    helper = BASE_DIR / 'libreoffice_export.py'

    result = subprocess.run(
        [
            '/usr/bin/python3',
            '-S',
            str(helper),
            str(excel_path),
            str(output_pdf),
        ],
        capture_output=True,
        text=True,
        timeout=240,
    )

    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(
            details or 'Excel to PDF conversion failed.'
        )

    if not output_pdf.exists():
        raise RuntimeError('Excel to PDF conversion produced no file.')


def create_final_pdf(letterhead_pdf, excel_pdf, final_pdf):
    letterhead_doc = fitz.open(str(letterhead_pdf))
    excel_doc = fitz.open(str(excel_pdf))
    output_doc = fitz.open()

    try:
        if len(letterhead_doc) <= LETTERHEAD_PAGE:
            raise ValueError('Letterhead PDF does not contain a first page.')

        if len(excel_doc) == 0:
            raise ValueError('Excel exported PDF contains no pages.')

        template_page = letterhead_doc[LETTERHEAD_PAGE]

        for page_number in range(len(excel_doc)):
            excel_page = excel_doc[page_number]
            crop_rect = get_non_white_bbox(excel_page)

            page = output_doc.new_page(
                width=cm_to_pt(A4_WIDTH_CM),
                height=cm_to_pt(A4_HEIGHT_CM),
            )

            page.show_pdf_page(
                page.rect,
                letterhead_doc,
                LETTERHEAD_PAGE,
                keep_proportion=False,
            )

            destination = calculate_destination_rect(
                page,
                crop_rect,
            )

            page.show_pdf_page(
                destination,
                excel_doc,
                page_number,
                keep_proportion=True,
                clip=crop_rect,
            )

        output_doc.save(
            str(final_pdf),
            garbage=4,
            deflate=True,
            clean=True,
        )

    finally:
        output_doc.close()
        excel_doc.close()
        letterhead_doc.close()


def extension_ok(name, allowed):
    return Path(name).suffix.lower() in allowed


@app.get('/')
def index():
    return render_template('index.html')


@app.get('/health')
def health():
    return jsonify({'status': 'ok'})


@app.post('/upload')
def upload():
    excel_file = request.files.get('excel_file')
    letterhead_file = request.files.get('letterhead')

    if not excel_file or not letterhead_file:
        return jsonify({'error': 'Please select both files.'}), 400

    if not excel_file.filename or not letterhead_file.filename:
        return jsonify({'error': 'Please select both files.'}), 400

    excel_name = secure_filename(excel_file.filename)
    letterhead_name = secure_filename(letterhead_file.filename)

    if not extension_ok(excel_name, ALLOWED_EXCEL):
        return jsonify({'error': 'Please upload an Excel file (.xlsx, .xls or .xlsm).'}), 400

    if not extension_ok(letterhead_name, ALLOWED_LETTERHEAD):
        return jsonify({'error': 'Please upload a PDF letterhead.'}), 400

    with CONVERSION_LOCK:
        with tempfile.TemporaryDirectory(prefix='excel_letterhead_') as temp_dir:
            temp_dir = Path(temp_dir)

            excel_path = temp_dir / excel_name
            letterhead_path = temp_dir / letterhead_name
            excel_pdf = temp_dir / 'excel_export.pdf'
            final_pdf = temp_dir / 'final_letterhead.pdf'

            excel_file.save(excel_path)
            letterhead_file.save(letterhead_path)

            try:
                export_excel_to_pdf(
                    excel_path,
                    excel_pdf,
                )

                create_final_pdf(
                    letterhead_path,
                    excel_pdf,
                    final_pdf,
                )

                pdf_bytes = final_pdf.read_bytes()

            except Exception as exc:
                app.logger.exception('PDF generation failed')
                return jsonify({
                    'error': str(exc),
                }), 500

    download_name = (
        Path(excel_name).stem
        + '_WITH_LETTERHEAD.pdf'
    )

    return send_file(
        BytesIO(pdf_bytes),
        as_attachment=True,
        download_name=download_name,
        mimetype='application/pdf',
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '5000'))
    app.run(
        host='0.0.0.0',
        port=port,
        debug=False,
    )
