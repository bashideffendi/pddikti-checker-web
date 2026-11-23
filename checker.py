"""
PDDIKTI Student Status Checker — Web Version
Refactored dari Pddikti_Checker.py untuk dipakai di Flask web app.

Catatan sumber data:
- App ini membaca data lewat Sivitas API (komunitas, bukan resmi PDDIKTI).
- API resmi PDDIKTI (api-pddikti.kemdiktisaintek.go.id) sekarang mewajibkan
  token reCAPTCHA dari browser, jadi tidak dipakai di sini.
- Kalau sumber data gagal, fungsi API me-raise SourceUnavailable. Jangan pernah
  diterjemahkan jadi "tidak ditemukan": itu dua hal berbeda.
"""

import math
import os
import re
import threading
import time
from urllib.parse import quote

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# KONFIGURASI
# ============================================================
_DEFAULT_BASES = [
    "https://pddikti.fastapicloud.dev/api",
    "https://pddikti.rone.dev/api",
]
# Bisa diganti lewat env, dipisah koma. Contoh: PDDIKTI_API_BASES=https://a/api,https://b/api
API_BASES = [
    b.strip().rstrip('/')
    for b in os.environ.get('PDDIKTI_API_BASES', ','.join(_DEFAULT_BASES)).split(',')
    if b.strip()
]
OFFICIAL_SITE = "https://pddikti.kemdiktisaintek.go.id"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept": "application/json",
}
REQUEST_TIMEOUT = 20
DELAY_SECONDS = 0.5           # jeda antar-request saat cek massal
MAX_COOLDOWN_SECONDS = 900    # sumber yang menolak (429/5xx) diistirahatkan maks. 15 menit
CONNECT_COOLDOWN_SECONDS = 60 # sumber yang tidak bisa dihubungi diistirahatkan 1 menit

# Kata umum nama kampus — tidak boleh dipakai sebagai bukti "PT cocok".
PT_GENERIC_WORDS = {
    "UNIVERSITAS", "UNIVERSITY", "INSTITUT", "INSTITUTE", "SEKOLAH", "TINGGI",
    "POLITEKNIK", "AKADEMI", "NEGERI", "SWASTA", "ISLAM", "ILMU", "TEKNOLOGI",
    "KESEHATAN", "EKONOMI", "KEGURUAN", "PENDIDIKAN", "AGAMA", "MUHAMMADIYAH",
    "NAHDLATUL", "ULAMA", "KRISTEN", "KATOLIK", "STIKES", "STIE", "STKIP",
    "STMIK", "STIT", "STAI", "STAIN", "IAIN", "UIN", "TEKNIK", "MANAJEMEN",
    "INFORMATIKA", "KOMPUTER", "BISNIS", "HUKUM", "SAINS", "DAN", "INDONESIA",
}


class SourceUnavailable(Exception):
    """Sumber data PDDIKTI tidak bisa dipakai saat ini (down, dibatasi, format berubah)."""

    def __init__(self, message, retry_after=None):
        super().__init__(message)
        self.message = message
        self.retry_after = retry_after


_session = requests.Session()
_session.headers.update(HEADERS)
_state_lock = threading.Lock()
_cooldown_until = {}   # base_url -> epoch detik
_last_problem = {}     # base_url -> teks masalah terakhir


# ============================================================
# FUNGSI UTILITAS
# ============================================================
def cell_text(value):
    """Nilai sel Excel -> teks bersih. NaN/None jadi ''. 12345.0 jadi '12345'."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    text = str(value).strip()
    if text.lower() in ("nan", "none", "nat"):
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def clean_nim(nim):
    return re.sub(r'[\s.\-]', '', cell_text(nim))


def normalize_query(text):
    """Rapikan input pencarian: buang gelar di belakang koma dan spasi ganda."""
    text = str(text or "").strip()
    if ',' in text:
        text = text.split(',', 1)[0]
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def _tokens(text):
    return [w for w in re.split(r'[^A-Z0-9]+', str(text).upper()) if w]


def nim_match(nim_excel, nim_pddikti):
    a = clean_nim(nim_excel).upper()
    b = clean_nim(nim_pddikti).upper()
    if not a or not b:
        return False
    if a == b:
        return True
    shorter = min(len(a), len(b))
    return shorter >= 6 and (a in b or b in a)


def nama_match(nama_excel, nama_pddikti):
    ta = _tokens(nama_excel)
    tb = _tokens(nama_pddikti)
    if not ta or not tb:
        return False
    if ta == tb:
        return True
    sa, sb = set(ta), set(tb)
    if sa <= sb or sb <= sa:
        return True
    words_a = {w for w in ta if len(w) >= 3}
    words_b = {w for w in tb if len(w) >= 3}
    return len(words_a & words_b) >= 2


def pt_match(pt_excel, record):
    """PT cocok kalau ada kata khas (bukan kata umum) yang sama, atau singkatan sama."""
    pt_tokens = set(_tokens(pt_excel))
    if not pt_tokens:
        return False
    r_pt = set(_tokens(record.get('nama_pt', '')))
    r_singkat = clean_nim(record.get('singkatan_pt', '')).upper()
    distinctive = {w for w in pt_tokens if len(w) > 2 and w not in PT_GENERIC_WORDS}
    if distinctive & r_pt:
        return True
    if r_singkat and len(r_singkat) >= 2 and r_singkat in pt_tokens:
        return True
    return False


# ============================================================
# FUNGSI API
# ============================================================
def _retry_after_seconds(resp):
    hint = None
    try:
        body = resp.json()
        if isinstance(body, dict):
            hint = body.get('retry_after_seconds')
    except ValueError:
        pass
    if hint is None:
        hint = resp.headers.get('Retry-After')
    try:
        return int(float(hint))
    except (TypeError, ValueError):
        return None


def _describe_status(code):
    if code == 429:
        return "terlalu banyak permintaan, HTTP 429"
    if code == 403:
        return "akses ditolak, HTTP 403"
    if code == 503:
        return "layanan sedang dibatasi, HTTP 503"
    if code >= 500:
        return f"server bermasalah, HTTP {code}"
    return f"respons tidak terduga, HTTP {code}"


def _get_json(path):
    """
    Ambil JSON dari sumber pertama yang menjawab normal.
    Return: payload JSON, atau None kalau sumber bilang data tidak ada (404 data).
    Raise: SourceUnavailable kalau tidak ada sumber yang bisa dipakai.
    """
    problems = []   # (alasan, detik tunggu) per sumber, urut prioritas
    for base in API_BASES:
        now = time.time()
        with _state_lock:
            until = _cooldown_until.get(base, 0)
            if until > now:
                problems.append((_last_problem.get(base, "sedang diistirahatkan"), until - now))
                continue
        try:
            resp = _session.get(f"{base}{path}", timeout=REQUEST_TIMEOUT)
        except requests.RequestException:
            with _state_lock:
                _cooldown_until[base] = now + CONNECT_COOLDOWN_SECONDS
                _last_problem[base] = "tidak bisa dihubungi"
            problems.append(("tidak bisa dihubungi", CONNECT_COOLDOWN_SECONDS))
            continue

        if resp.status_code == 200:
            try:
                data = resp.json()
            except ValueError:
                problems.append(("respons bukan JSON", None))
                continue
            if isinstance(data, dict) and data.get('success') is False:
                code = (data.get('status') or {}).get('http_code', 503)
                retry = _retry_after_seconds(resp) or 300
                with _state_lock:
                    _cooldown_until[base] = now + min(retry, MAX_COOLDOWN_SECONDS)
                    _last_problem[base] = _describe_status(int(code))
                problems.append((_describe_status(int(code)), min(retry, MAX_COOLDOWN_SECONDS)))
                continue
            with _state_lock:
                _cooldown_until.pop(base, None)
                _last_problem.pop(base, None)
            return data

        if resp.status_code == 404:
            # FastAPI mengembalikan {"detail":"Not Found"} kalau endpoint-nya tidak ada.
            # Itu tanda API berubah, bukan tanda mahasiswa tidak terdaftar.
            try:
                body = resp.json()
            except ValueError:
                body = None
            if isinstance(body, dict) and body.get('detail') == 'Not Found':
                problems.append(("endpoint tidak ditemukan, API mungkin berubah", None))
                continue
            return None

        retry = _retry_after_seconds(resp) or 300
        reason = _describe_status(resp.status_code)
        with _state_lock:
            _cooldown_until[base] = now + min(retry, MAX_COOLDOWN_SECONDS)
            _last_problem[base] = reason
        problems.append((reason, min(retry, MAX_COOLDOWN_SECONDS)))

    # Laporkan masalah sumber utama (paling informatif), bukan yang tercepat pulih.
    reason, wait = problems[0] if problems else ("tidak ada sumber data yang dikonfigurasi", None)
    raise SourceUnavailable(
        f"Sumber data PDDIKTI tidak bisa dipakai ({reason}).",
        retry_after=int(wait) if wait else None,
    )


def _unwrap_list(data):
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ('data', 'mahasiswa', 'results', 'result', 'items'):
            value = data.get(key)
            if isinstance(value, list):
                return value
            if isinstance(value, dict):
                inner = _unwrap_list(value)
                if inner is not None:
                    return inner
    return None


def _normalize_record(r):
    return {
        'id': cell_text(r.get('id')),
        'nama': cell_text(r.get('nama')),
        'nim': cell_text(r.get('nim')),
        'nama_pt': cell_text(r.get('nama_pt')),
        'singkatan_pt': cell_text(r.get('sinkatan_pt') or r.get('singkatan_pt')),
        'nama_prodi': cell_text(r.get('nama_prodi') or r.get('prodi')),
    }


def api_search(keyword):
    """Cari mahasiswa. Return list (bisa kosong = memang tidak ada). Raise SourceUnavailable."""
    keyword = normalize_query(keyword)
    if not keyword:
        return []
    data = _get_json(f"/search/mhs/{quote(keyword, safe='')}/")
    if data is None:
        return []
    rows = _unwrap_list(data)
    if rows is None:
        if isinstance(data, dict) and not data:
            return []
        raise SourceUnavailable("Format respons sumber data berubah sehingga hasil tidak bisa dibaca.")
    return [_normalize_record(r) for r in rows if isinstance(r, dict)]


def api_detail(id_mhs):
    """Detail mahasiswa. Return dict atau None (tidak ada). Raise SourceUnavailable."""
    id_mhs = str(id_mhs or "").strip()
    if not id_mhs:
        return None
    data = _get_json(f"/mhs/detail/{quote(id_mhs, safe='')}/")
    if data is None:
        return None
    if isinstance(data, dict) and isinstance(data.get('data'), dict):
        data = data['data']
    if isinstance(data, list):
        data = data[0] if data and isinstance(data[0], dict) else None
    if not isinstance(data, dict):
        raise SourceUnavailable("Format respons sumber data berubah sehingga detail tidak bisa dibaca.")
    return data


def source_status():
    """Ringkasan kondisi sumber data untuk ditampilkan di UI (tanpa request baru)."""
    now = time.time()
    with _state_lock:
        blocked = {b: _cooldown_until.get(b, 0) for b in API_BASES}
        problems = dict(_last_problem)
    if any(until <= now for until in blocked.values()):
        return {'ok': True}   # minimal satu sumber boleh dicoba
    first = API_BASES[0]
    return {
        'ok': False,
        'reason': problems.get(first, 'sedang bermasalah'),
        'retry_after': int(blocked[first] - now),
    }


def probe_source():
    """Tes ringan ke sumber data. Return dict status."""
    try:
        api_search("andi")
        return {'ok': True}
    except SourceUnavailable as e:
        return {'ok': False, 'reason': e.message, 'retry_after': e.retry_after}


# ============================================================
# LOGIKA PENCARIAN & PENCOCOKAN
# ============================================================
def find_best_match(results, nim, nama, pt):
    nim_c = clean_nim(nim).upper()

    if nim_c:
        for r in results:
            if clean_nim(r.get('nim', '')).upper() == nim_c:
                return r, "NIM cocok persis"
        for r in results:
            if nim_match(nim, r.get('nim', '')):
                return r, "NIM cocok sebagian"

    if pt:
        for r in results:
            if nama_match(nama, r.get('nama', '')) and pt_match(pt, r):
                return r, "nama + PT cocok"

    nama_up = ' '.join(_tokens(nama))
    exact = [r for r in results if ' '.join(_tokens(r.get('nama', ''))) == nama_up]
    if nama_up and len(exact) == 1 and not pt:
        return exact[0], "nama unik cocok (tanpa PT)"

    return None, None


def search_student(nim, nama, pt):
    """Return (record|None, keterangan). Raise SourceUnavailable."""
    nim_cleaned = clean_nim(nim)
    last_results = []

    if nim_cleaned:
        last_results = api_search(nim_cleaned)
        match, method = find_best_match(last_results, nim, nama, pt)
        if match:
            return match, f"via NIM: {method}"
        time.sleep(DELAY_SECONDS)

    if nama:
        results = api_search(nama)
        if results:
            last_results = results
        match, method = find_best_match(results, nim, nama, pt)
        if match:
            return match, f"via nama: {method}"

    if last_results:
        return None, f"Ada {len(last_results)} hasil, tetapi NIM/PT tidak cocok"
    return None, "Tidak ditemukan di PDDIKTI"


# ============================================================
# DETEKSI KOLOM
# ============================================================
def detect_columns(df):
    cols = df.columns.tolist()
    col_map = {"nim": None, "nama": None, "pt": None}

    for c in cols:
        c_lower = str(c).lower()
        if col_map["nim"] is None and ("nim" in c_lower or "nip" in c_lower):
            col_map["nim"] = c
        elif col_map["pt"] is None and any(x in c_lower for x in ["perguruan", "universitas", "kampus", "nama pt"]):
            col_map["pt"] = c
        elif col_map["pt"] is None and re.search(r'\bpt\b', c_lower):
            col_map["pt"] = c
        elif col_map["nama"] is None and "nama" in c_lower:
            col_map["nama"] = c

    if not col_map["nim"] and len(cols) > 1:
        col_map["nim"] = cols[1]
    if not col_map["nama"] and len(cols) > 2:
        col_map["nama"] = cols[2]
    if not col_map["pt"] and len(cols) > 3:
        col_map["pt"] = cols[3]

    return col_map


def is_sub_header(row):
    values = [cell_text(v) for v in row.values]
    if all(v == '' or (v.isdigit() and len(v) <= 2) for v in values):
        return True
    if len(values) >= 2 and values[0].isdigit() and values[1].isdigit() and len(values[1]) <= 2:
        return True
    return False


# ============================================================
# FORMAT EXCEL OUTPUT
# ============================================================
NEW_COLUMNS = ["STATUS MAHASISWA (PDDIKTI)", "PT (PDDIKTI)", "PRODI (PDDIKTI)",
               "NIM (PDDIKTI)", "JENJANG", "KETERANGAN"]


STATUS_COLORS = {            # (warna huruf, warna isi)
    "green": ("12723B", "EBF6EF"),
    "blue": ("1C5BC4", "EDF3FD"),
    "amber": ("9A5A06", "FCF5E6"),
    "red": ("B42318", "FDF0EF"),
}


def _status_tone(value):
    s = cell_text(value).lower()
    if not s:
        return None
    if "aktif" in s and "non" not in s:
        return "green"
    if "lulus" in s:
        return "blue"
    if "cuti" in s:
        return "amber"
    return "red"


def format_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb.active

    total_cols = ws.max_column
    total_rows = ws.max_row
    new_col_start = total_cols - len(NEW_COLUMNS) + 1

    header_fill_orig = PatternFill(start_color="1F2A3A", end_color="1F2A3A", fill_type="solid")
    header_fill_new = PatternFill(start_color="3D4B5E", end_color="3D4B5E", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    data_font = Font(size=10, name="Arial")
    thin_border = Border(
        left=Side(style='thin', color="C9C4B8"), right=Side(style='thin', color="C9C4B8"),
        top=Side(style='thin', color="C9C4B8"), bottom=Side(style='thin', color="C9C4B8")
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    green_fill = PatternFill(start_color="DCEBDD", end_color="DCEBDD", fill_type="solid")
    red_fill = PatternFill(start_color="F3D9D4", end_color="F3D9D4", fill_type="solid")
    yellow_fill = PatternFill(start_color="F4E8C8", end_color="F4E8C8", fill_type="solid")
    gray_fill = PatternFill(start_color="E3E1DC", end_color="E3E1DC", fill_type="solid")
    stripe_fill = PatternFill(start_color="F6F4EF", end_color="F6F4EF", fill_type="solid")

    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill_orig if col_idx < new_col_start else header_fill_new
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    ket_col = total_cols

    for row_idx in range(2, total_rows + 1):
        is_odd = (row_idx % 2 == 1)
        ket_val = str(ws.cell(row=row_idx, column=ket_col).value or "")
        row_color = None
        if ket_val.startswith("Ditemukan"):
            row_color = green_fill
        elif ket_val.startswith("Tidak ditemukan"):
            row_color = red_fill
        elif "tidak cocok" in ket_val.lower() or ket_val.startswith("Cocok"):
            row_color = yellow_fill
        elif ket_val.startswith(("Belum dicek", "Gagal cek", "Dilewati")):
            row_color = gray_fill

        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_align if col_idx <= 2 else left_align
            if col_idx >= new_col_start and row_color:
                cell.fill = row_color
            elif is_odd:
                cell.fill = stripe_fill

        # Kolom status diberi warna sendiri: aktif hijau, lulus biru, cuti kuning, lainnya merah.
        status_cell = ws.cell(row=row_idx, column=new_col_start)
        tone = _status_tone(status_cell.value)
        if tone:
            font_color, fill_color = STATUS_COLORS[tone]
            status_cell.font = Font(size=10, name="Arial", bold=True, color=font_color)
            status_cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    for col_idx in range(1, new_col_start):
        header_val = str(ws.cell(row=1, column=col_idx).value or "")
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = max(len(header_val) + 4, 12)

    new_widths = [30, 28, 22, 18, 13, 42]
    for i, w in enumerate(new_widths):
        col_letter = ws.cell(row=1, column=new_col_start + i).column_letter
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=total_cols).column_letter}{total_rows}"

    wb.save(filepath)


# ============================================================
# MAIN RUNNER (dengan callback progress)
# ============================================================
def _lower_first(text):
    """Huruf awal jadi kecil (dipakai setelah titik dua), kecuali singkatan seperti NIM."""
    text = str(text or "")
    if not text or text.split(" ", 1)[0].isupper():
        return text
    return text[0].lower() + text[1:]


def _empty_result(keterangan):
    return {
        "status_saat_ini": "",
        "nama_pt_pddikti": "",
        "prodi_pddikti": "",
        "nim_pddikti": "",
        "jenjang": "",
        "keterangan": keterangan,
    }


def run_checker(input_file, output_file, on_progress=None):
    """
    Jalankan pengecekan PDDIKTI.

    Kalau sumber data mati di tengah jalan, proses berhenti, baris yang sudah
    dicek tetap disimpan, sisanya ditandai "Belum dicek".

    Args:
        input_file:   Path file Excel input
        output_file:  Path file Excel output
        on_progress:  Callback function(current, total, nama, status_text)
    """
    df = pd.read_excel(input_file, dtype=str)

    skip_header_row = len(df) > 0 and is_sub_header(df.iloc[0])
    if skip_header_row:
        df = df.iloc[1:].reset_index(drop=True)

    col_map = detect_columns(df)
    if not col_map['nama'] and not col_map['nim']:
        raise ValueError("Kolom NIM dan NAMA MAHASISWA tidak ditemukan di file.")

    total = len(df)
    results_data = []
    seen = {}
    aborted = None

    for idx, row in df.iterrows():
        nama = cell_text(row[col_map['nama']]) if col_map['nama'] else ""
        nim = cell_text(row[col_map['nim']]) if col_map['nim'] else ""
        pt = cell_text(row[col_map['pt']]) if col_map['pt'] else ""
        label = nama or nim or f"Baris {idx + 2}"

        if not nama and not nim:
            results_data.append(_empty_result("Dilewati: NIM dan nama kosong"))
            if on_progress:
                on_progress(idx + 1, total, label, "Dilewati (baris kosong)")
            continue

        key = (clean_nim(nim).upper(), ' '.join(_tokens(nama)), ' '.join(_tokens(pt)))
        if key in seen:
            results_data.append(dict(seen[key]))
            if on_progress:
                on_progress(idx + 1, total, label, f"Duplikat: {_lower_first(seen[key]['keterangan'])}")
            continue

        if on_progress:
            on_progress(idx + 1, total, label, "Mencari...")

        try:
            matched, method = search_student(nim, nama, pt)
            if matched:
                time.sleep(DELAY_SECONDS)
                detail = api_detail(matched['id'])
                if detail:
                    result_info = {
                        "status_saat_ini": cell_text(detail.get('status_saat_ini')),
                        "nama_pt_pddikti": cell_text(detail.get('nama_pt')),
                        "prodi_pddikti": cell_text(detail.get('prodi')),
                        "nim_pddikti": cell_text(detail.get('nim')),
                        "jenjang": cell_text(detail.get('jenjang')),
                        "keterangan": f"Ditemukan ({method})",
                    }
                else:
                    result_info = _empty_result(f"Cocok ({method}), tetapi detail tidak tersedia")
                    result_info["nama_pt_pddikti"] = matched.get('nama_pt', '')
                    result_info["prodi_pddikti"] = matched.get('nama_prodi', '')
                    result_info["nim_pddikti"] = matched.get('nim', '')
            else:
                result_info = _empty_result(method or "Tidak ditemukan di PDDIKTI")
        except SourceUnavailable as e:
            aborted = e.message
            results_data.append(_empty_result(f"Gagal cek: {_lower_first(e.message)}"))
            if on_progress:
                on_progress(idx + 1, total, label, f"Gagal cek: {_lower_first(e.message)}")
            break

        seen[key] = result_info
        results_data.append(result_info)

        if on_progress:
            on_progress(idx + 1, total, label, result_info["status_saat_ini"] or result_info["keterangan"])

        time.sleep(DELAY_SECONDS)

    # Baris yang belum sempat dicek (karena berhenti)
    while len(results_data) < total:
        results_data.append(_empty_result("Belum dicek (proses berhenti karena sumber data tidak tersedia)"))

    # Tulis Excel output
    df_out = pd.read_excel(input_file)
    padding = [_empty_result("")] if skip_header_row else []
    rows = padding + results_data
    rows += [_empty_result("")] * (len(df_out) - len(rows))
    rows = rows[:len(df_out)]

    field_by_col = dict(zip(NEW_COLUMNS, ["status_saat_ini", "nama_pt_pddikti", "prodi_pddikti",
                                          "nim_pddikti", "jenjang", "keterangan"]))
    for col, field in field_by_col.items():
        df_out[col] = [r.get(field, "") for r in rows]

    df_out.to_excel(output_file, index=False)
    format_excel(output_file)

    def count(pred):
        return sum(1 for r in results_data if pred(r.get("keterangan", "")))

    return {
        "total": total,
        "found": count(lambda k: k.startswith("Ditemukan")),
        "not_found": count(lambda k: k.startswith("Tidak ditemukan")),
        "no_match": count(lambda k: "tidak cocok" in k.lower() or k.startswith("Cocok")),
        "unchecked": count(lambda k: k.startswith(("Gagal cek", "Belum dicek"))),
        "skipped": count(lambda k: k.startswith("Dilewati")),
        "aborted": aborted,
    }
