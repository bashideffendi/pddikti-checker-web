"""
PDDIKTI Checker — Flask Web App
"""

import os
import re
import time
import uuid
import threading
from datetime import datetime
from urllib.parse import quote

from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

from checker import (
    run_checker, api_search, api_detail, normalize_query,
    source_status, probe_source, SourceUnavailable, OFFICIAL_SITE,
)

app = Flask(__name__, static_folder='static', static_url_path='/static')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
RESULT_DIR = os.path.join(BASE_DIR, 'results')
FILE_TTL_SECONDS = 24 * 3600   # file upload & hasil dihapus setelah 24 jam
JOB_ID_RE = re.compile(r'^(?:[a-f0-9]{32}|[a-f0-9]{8})$')  # 8 = format lama di riwayat

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# In-memory job tracking (gunicorn 1 worker)
jobs = {}
jobs_lock = threading.Lock()
_probe_cache = {'at': 0, 'value': None}


def manual_url(q=""):
    """Tautan cek manual di situs resmi PDDIKTI."""
    q = normalize_query(q)
    return f"{OFFICIAL_SITE}/search/{quote(q, safe='')}" if q else OFFICIAL_SITE


def unavailable_response(e, q=""):
    return jsonify({
        'error': e.message,
        'code': 'SOURCE_UNAVAILABLE',
        'retry_after': e.retry_after,
        'manual_url': manual_url(q),
    }), 503


def cleanup_old_files():
    cutoff = time.time() - FILE_TTL_SECONDS
    for folder in (UPLOAD_DIR, RESULT_DIR):
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            try:
                if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                    os.remove(path)
            except OSError:
                pass
    with jobs_lock:
        for job_id in [j for j, v in jobs.items() if v.get('created_ts', 0) < cutoff]:
            jobs.pop(job_id, None)


def process_job(job_id, input_path, output_path):
    """Background worker for processing a job."""
    job = jobs[job_id]
    try:
        job['status'] = 'processing'

        def on_progress(current, total, nama, status_text):
            job['current'] = current
            job['total'] = total
            job['seq'] += 1
            job['logs'].append({
                'seq': job['seq'],
                'index': current,
                'total': total,
                'nama': nama,
                'status': status_text,
            })
            if len(job['logs']) > 300:
                del job['logs'][:-300]

        summary = run_checker(input_path, output_path, on_progress=on_progress)
        job['summary'] = summary
        job['status'] = 'done'
    except ValueError as e:
        job['error'] = str(e)
        job['status'] = 'error'
    except Exception:
        app.logger.exception("Job %s gagal", job_id)
        job['error'] = 'File tidak bisa diproses. Pastikan formatnya sesuai template.'
        job['status'] = 'error'
    finally:
        try:
            os.remove(input_path)
        except OSError:
            pass


# ============================================================
# ROUTES
# ============================================================

@app.route('/')
def index():
    return render_template('index.html', official_site=OFFICIAL_SITE)


@app.route('/status')
def status():
    """Kondisi sumber data. Pakai status terakhir; probe ulang maks. tiap 2 menit."""
    current = source_status()
    if not current['ok']:
        return jsonify(current)
    now = time.time()
    if _probe_cache['value'] is None or now - _probe_cache['at'] > 120:
        _probe_cache['value'] = probe_source()
        _probe_cache['at'] = now
    return jsonify(_probe_cache['value'])


@app.route('/search')
def search():
    """Quick search by name or NIM — returns list of candidates."""
    q = normalize_query(request.args.get('q', ''))
    if len(q) < 3:
        return jsonify({'error': 'Ketik minimal tiga karakter nama atau NIM.'}), 400
    if len(q) > 100:
        return jsonify({'error': 'Kata kunci terlalu panjang (maksimal 100 karakter).'}), 400
    try:
        results = api_search(q)
    except SourceUnavailable as e:
        return unavailable_response(e, q)
    except Exception:
        app.logger.exception("Search gagal")
        return jsonify({'error': 'Pencarian gagal karena kesalahan di server.'}), 500
    _probe_cache['value'] = {'ok': True}
    _probe_cache['at'] = time.time()
    return jsonify({
        'query': q,
        'results': results[:50],
        'total': len(results),
        'manual_url': manual_url(q),
    })


@app.route('/detail')
def detail():
    """Get full detail of a single student by id."""
    id_mhs = request.args.get('id', '').strip()
    if not id_mhs or len(id_mhs) > 200:
        return jsonify({'error': 'ID mahasiswa tidak valid.'}), 400
    try:
        data = api_detail(id_mhs)
    except SourceUnavailable as e:
        return unavailable_response(e)
    except Exception:
        app.logger.exception("Detail gagal")
        return jsonify({'error': 'Detail gagal dimuat karena kesalahan di server.'}), 500
    if not data:
        return jsonify({'error': 'Detail mahasiswa tidak tersedia di sumber data.'}), 404
    return jsonify(data)


@app.route('/template')
def download_template():
    template_path = os.path.join(BASE_DIR, 'static', 'template_data_mahasiswa.xlsx')
    return send_file(template_path, as_attachment=True, download_name='Template_Data_Mahasiswa.xlsx')


@app.route('/upload', methods=['POST'])
def upload():
    cleanup_old_files()

    file = request.files.get('file')
    if not file or file.filename == '':
        return jsonify({'error': 'Pilih file Excel terlebih dahulu.'}), 400
    if not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'File harus berformat .xlsx.'}), 400
    head = file.stream.read(4)
    file.stream.seek(0)
    if head != b'PK\x03\x04':
        return jsonify({'error': 'Isi file bukan format .xlsx yang valid.'}), 400

    job_id = uuid.uuid4().hex
    original_filename = secure_filename(file.filename) or 'data.xlsx'

    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_{original_filename}")
    output_path = os.path.join(RESULT_DIR, f"{job_id}_Hasil_{original_filename}")
    file.save(input_path)

    with jobs_lock:
        jobs[job_id] = {
            'status': 'queued',
            'current': 0,
            'total': 0,
            'seq': 0,
            'logs': [],
            'summary': None,
            'error': None,
            'filename': original_filename,
            'output_path': output_path,
            'created': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'created_ts': time.time(),
        }

    threading.Thread(target=process_job, args=(job_id, input_path, output_path), daemon=True).start()
    return jsonify({'job_id': job_id})


@app.route('/progress/<job_id>')
def progress(job_id):
    """Polling progres. ?since=<seq> mengembalikan log setelah nomor itu."""
    if not JOB_ID_RE.match(job_id):
        return jsonify({'error': 'ID proses tidak valid.'}), 400
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Proses tidak ditemukan. Server mungkin baru dimulai ulang. Silakan upload ulang file.'}), 404
    try:
        since = int(request.args.get('since', 0))
    except ValueError:
        since = 0
    payload = {
        'status': job['status'],
        'current': job['current'],
        'total': job['total'],
        'seq': job['seq'],
        'logs': [log for log in job['logs'] if log['seq'] > since],
    }
    if job['status'] == 'done':
        payload['summary'] = job['summary']
    elif job['status'] == 'error':
        payload['error'] = job['error']
    return jsonify(payload)


@app.route('/download/<job_id>')
def download(job_id):
    if not JOB_ID_RE.match(job_id):
        return jsonify({'error': 'ID proses tidak valid.'}), 400

    job = jobs.get(job_id)
    if job and job['status'] == 'done' and os.path.exists(job['output_path']):
        return send_file(job['output_path'], as_attachment=True,
                         download_name=f"Hasil_{job['filename']}")

    prefix = f"{job_id}_Hasil_"
    for name in os.listdir(RESULT_DIR):
        if name.startswith(prefix):
            return send_file(os.path.join(RESULT_DIR, name), as_attachment=True,
                             download_name=name[len(job_id) + 1:])

    return jsonify({'error': 'File hasil sudah tidak tersedia karena dihapus setelah 24 jam atau server dimulai ulang. Silakan proses ulang.'}), 404


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
