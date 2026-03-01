import os
import uuid
import subprocess
import shutil
import time
from flask import (
    Flask, render_template, request, jsonify,
    send_file, after_this_request
)
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max
app.config['UPLOAD_FOLDER'] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'uploads'
)
app.config['COMPRESSED_FOLDER'] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'compressed'
)

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['COMPRESSED_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}

# Ghostscript compression settings mapped to quality levels
# /screen   = 72 dpi  — smallest, lowest quality
# /ebook    = 150 dpi — medium quality
# /printer  = 300 dpi — high quality
# /prepress = 300 dpi — highest quality, color preserving
COMPRESSION_LEVELS = {
    'extreme': {
        'gs_setting': '/screen',
        'label': 'Extreme Compression',
        'description': '72 DPI — Smallest file size, lower image quality',
        'dpi': 72
    },
    'high': {
        'gs_setting': '/ebook',
        'label': 'High Compression',
        'description': '150 DPI — Good balance of size and quality',
        'dpi': 150
    },
    'medium': {
        'gs_setting': '/printer',
        'label': 'Medium Compression',
        'description': '300 DPI — Minimal quality loss, moderate reduction',
        'dpi': 300
    },
    'low': {
        'gs_setting': '/prepress',
        'label': 'Low Compression',
        'description': '300 DPI — Best quality, least compression',
        'dpi': 300
    }
}


def allowed_file(filename):
    return (
        '.' in filename
        and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def find_ghostscript():
    """Find the Ghostscript executable on the system."""
    # Common Ghostscript executable names
    gs_names = ['gs', 'gswin64c', 'gswin32c', 'gswin64', 'gswin32']

    for name in gs_names:
        path = shutil.which(name)
        if path:
            return path

    # Check common installation paths on Windows
    common_paths = [
        r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs9.55.0\bin\gswin64c.exe',
        r'C:\Program Files\gs\gs9.54.0\bin\gswin64c.exe',
        r'C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe',
    ]

    for path in common_paths:
        if os.path.exists(path):
            return path

    return None


def compress_pdf_with_ghostscript(
    input_path, output_path, level='high', custom_dpi=None
):
    """
    Compress a PDF file using Ghostscript.

    Args:
        input_path: Path to the input PDF file
        output_path: Path where the compressed PDF will be saved
        level: Compression level ('extreme', 'high', 'medium', 'low')
        custom_dpi: Optional custom DPI value (overrides level default)

    Returns:
        dict with compression results or raises an exception
    """
    gs_path = find_ghostscript()
    if not gs_path:
        raise RuntimeError(
            'Ghostscript is not installed or not found in PATH. '
            'Please install Ghostscript: https://www.ghostscript.com/releases/gsdnld.html'
        )

    if level not in COMPRESSION_LEVELS:
        level = 'high'

    settings = COMPRESSION_LEVELS[level]
    gs_quality = settings['gs_setting']
    dpi = custom_dpi if custom_dpi else settings['dpi']

    # Build Ghostscript command
    gs_command = [
        gs_path,
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        f'-dPDFSETTINGS={gs_quality}',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        '-dDetectDuplicateImages=true',
        '-dCompressFonts=true',
        '-dSubsetFonts=true',
        f'-dDownsampleColorImages=true',
        f'-dColorImageResolution={dpi}',
        f'-dDownsampleGrayImages=true',
        f'-dGrayImageResolution={dpi}',
        f'-dDownsampleMonoImages=true',
        f'-dMonoImageResolution={dpi}',
        '-dAutoRotatePages=/None',
        '-dColorImageDownsampleType=/Bicubic',
        '-dGrayImageDownsampleType=/Bicubic',
        f'-sOutputFile={output_path}',
        input_path
    ]

    try:
        result = subprocess.run(
            gs_command,
            capture_output=True,
            text=True,
            timeout=300  # 5 minute timeout
        )

        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or 'Unknown Ghostscript error'
            raise RuntimeError(f'Ghostscript compression failed: {error_msg}')

        # Get file sizes for comparison
        original_size = os.path.getsize(input_path)
        compressed_size = os.path.getsize(output_path)

        # If compressed file is larger, copy original
        if compressed_size >= original_size:
            shutil.copy2(input_path, output_path)
            compressed_size = original_size

        reduction_percent = (
            ((original_size - compressed_size) / original_size) * 100
            if original_size > 0 else 0
        )

        return {
            'success': True,
            'original_size': original_size,
            'compressed_size': compressed_size,
            'reduction_percent': round(reduction_percent, 1),
            'level': level,
            'dpi': dpi
        }

    except subprocess.TimeoutExpired:
        raise RuntimeError(
            'Compression timed out. The PDF file may be too large or complex.'
        )
    except FileNotFoundError:
        raise RuntimeError(
            'Ghostscript executable not found. Please verify installation.'
        )


def format_file_size(size_bytes):
    """Format bytes into human-readable file size."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def cleanup_old_files(folder, max_age_seconds=3600):
    """Remove files older than max_age_seconds from the given folder."""
    now = time.time()
    try:
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)
            if os.path.isfile(filepath):
                file_age = now - os.path.getmtime(filepath)
                if file_age > max_age_seconds:
                    os.remove(filepath)
    except Exception:
        pass


# ──────────────── Routes ────────────────

@app.route('/')
def index():
    return render_template('compress.html')


@app.route('/compress-pdf', methods=['GET'])
def compress_page():
    return render_template('compress.html')


@app.route('/api/compress', methods=['POST'])
def api_compress():
    """API endpoint to compress a PDF file."""
    # Clean up old files first
    cleanup_old_files(app.config['UPLOAD_FOLDER'])
    cleanup_old_files(app.config['COMPRESSED_FOLDER'])

    # Validate file presence
    if 'pdf_file' not in request.files:
        return jsonify({
            'success': False,
            'error': 'No PDF file provided. Please select a file to compress.'
        }), 400

    file = request.files['pdf_file']

    if file.filename == '':
        return jsonify({
            'success': False,
            'error': 'No file selected. Please choose a PDF file.'
        }), 400

    if not allowed_file(file.filename):
        return jsonify({
            'success': False,
            'error': 'Invalid file type. Only PDF files are accepted.'
        }), 400

    # Get compression level
    level = request.form.get('level', 'high')
    if level not in COMPRESSION_LEVELS:
        level = 'high'

    # Get optional custom DPI
    custom_dpi = request.form.get('custom_dpi', None)
    if custom_dpi:
        try:
            custom_dpi = int(custom_dpi)
            if custom_dpi < 10 or custom_dpi > 1200:
                custom_dpi = None
        except (ValueError, TypeError):
            custom_dpi = None

    try:
        # Generate unique filenames
        unique_id = str(uuid.uuid4())
        original_filename = secure_filename(file.filename)
        name_without_ext = os.path.splitext(original_filename)[0]

        input_filename = f"{unique_id}_input.pdf"
        output_filename = f"{unique_id}_compressed.pdf"

        input_path = os.path.join(
            app.config['UPLOAD_FOLDER'], input_filename
        )
        output_path = os.path.join(
            app.config['COMPRESSED_FOLDER'], output_filename
        )

        # Save uploaded file
        file.save(input_path)

        # Verify it's a valid PDF
        with open(input_path, 'rb') as f:
            header = f.read(5)
            if header != b'%PDF-':
                os.remove(input_path)
                return jsonify({
                    'success': False,
                    'error': 'The uploaded file is not a valid PDF document.'
                }), 400

        # Compress using Ghostscript
        result = compress_pdf_with_ghostscript(
            input_path, output_path, level, custom_dpi
        )

        # Clean up the uploaded original
        try:
            os.remove(input_path)
        except Exception:
            pass

        # Build response
        return jsonify({
            'success': True,
            'download_id': unique_id,
            'original_filename': original_filename,
            'suggested_filename': f"{name_without_ext}_compressed.pdf",
            'original_size': result['original_size'],
            'compressed_size': result['compressed_size'],
            'original_size_formatted': format_file_size(
                result['original_size']
            ),
            'compressed_size_formatted': format_file_size(
                result['compressed_size']
            ),
            'reduction_percent': result['reduction_percent'],
            'level': result['level'],
            'level_label': COMPRESSION_LEVELS[result['level']]['label'],
            'dpi': result['dpi']
        })

    except RuntimeError as e:
        # Clean up files on error
        for path in [input_path, output_path]:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'An unexpected error occurred: {str(e)}'
        }), 500


@app.route('/api/download/<download_id>', methods=['GET'])
def api_download(download_id):
    """Download a compressed PDF file."""
    # Sanitize download_id
    try:
        uuid.UUID(download_id)
    except ValueError:
        return jsonify({'success': False, 'error': 'Invalid download ID.'}), 400

    output_filename = f"{download_id}_compressed.pdf"
    output_path = os.path.join(
        app.config['COMPRESSED_FOLDER'], output_filename
    )

    if not os.path.exists(output_path):
        return jsonify({
            'success': False,
            'error': 'File not found. It may have expired. Please compress again.'
        }), 404

    # Get custom filename from query parameter
    custom_name = request.args.get('filename', 'compressed.pdf')
    if not custom_name.lower().endswith('.pdf'):
        custom_name += '.pdf'

    @after_this_request
    def cleanup(response):
        """Remove the compressed file after sending it."""
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except Exception:
            pass
        return response

    return send_file(
        output_path,
        as_attachment=True,
        download_name=custom_name,
        mimetype='application/pdf'
    )


@app.route('/api/check-ghostscript', methods=['GET'])
def check_ghostscript():
    """Check if Ghostscript is installed and accessible."""
    gs_path = find_ghostscript()
    if gs_path:
        try:
            result = subprocess.run(
                [gs_path, '--version'],
                capture_output=True, text=True, timeout=10
            )
            version = result.stdout.strip() if result.returncode == 0 else 'Unknown'
            return jsonify({
                'installed': True,
                'path': gs_path,
                'version': version
            })
        except Exception:
            return jsonify({
                'installed': True,
                'path': gs_path,
                'version': 'Unknown'
            })
    else:
        return jsonify({
            'installed': False,
            'error': 'Ghostscript not found. Please install it.'
        })


# ──────────────── Run ────────────────

if __name__ == '__main__':
    # Verify Ghostscript on startup
    gs = find_ghostscript()
    if gs:
        print(f"✅ Ghostscript found at: {gs}")
    else:
        print("⚠️  WARNING: Ghostscript not found!")
        print("   Install it from: https://www.ghostscript.com/releases/gsdnld.html")
        print("   Ubuntu/Debian:  sudo apt-get install ghostscript")
        print("   macOS:          brew install ghostscript")
        print("   Windows:        Download from the link above")

    app.run(debug=True, host='0.0.0.0', port=5000)