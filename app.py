#!/usr/bin/env python3
"""
PEMILU OSIS — e-Voting Sekolah (Pemilihan Ketua OSIS)
Stack: Python 3 standard library (http.server + sqlite3) — tanpa
framework web. Satu-satunya dependensi eksternal: openpyxl, khusus
untuk ekspor Excel (.xlsx) laporan hasil akhir.

Tiga sisi aplikasi:
1. ADMIN (satu akun admin, wajib login): mengelola kandidat & daftar
   pemilih, mengatur jadwal voting dan tampilan hasil, memantau
   dashboard hasil langsung, mengunduh CSV, serta Backup & Pulihkan.
   Tema GELAP: navy tua + aksen toska & emas.
2. PEMILIH / VOTER (tanpa akun; cukup NIS): halaman terang ramah HP.
   Siswa memasukkan NIS -> terverifikasi -> surat suara -> konfirmasi
   -> kirim. Satu NIS hanya bisa memakai hak pilihnya SATU kali.
3. LAYAR TAMPIL (/layar, publik): papan hasil layar penuh untuk
   proyektor/TV — hanya angka agregat, TANPA identitas pemilih.

KERAHASIAAN SUARA (syarat mutlak): tabel `suara` HANYA menyimpan
kandidat_id + waktu. TIDAK ADA kolom/referensi pemilih di tabel itu,
jadi pilihan seorang siswa TIDAK PERNAH bisa dikaitkan dengan orangnya.
Admin hanya melihat partisipasi (siapa sudah/belum memilih), bukan
pilihan siapa pun.

Jalankan:  python3 app.py            (default http://127.0.0.1:8000)
           PORT=8080 python3 app.py
           VOTE_DB=/path/vote.db VOTE_UPLOADS=/path/uploads python3 app.py
"""
import csv
import datetime
import hashlib
import html
import io
import json
import os
import re
import secrets
import shutil
import sqlite3
import tempfile
import urllib.parse
import zipfile
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from zoneinfo import ZoneInfo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("VOTE_DB", os.path.join(BASE_DIR, "data", "vote.db"))
UPLOADS_DIR = os.environ.get("VOTE_UPLOADS", os.path.join(BASE_DIR, "data", "uploads"))
PORT = int(os.environ.get("PORT", "8000"))
TZ_JKT = ZoneInfo("Asia/Jakarta")

ALLOWED_EXT = {".jpg", ".jpeg", ".png", ".webp"}
MAX_FOTO_BYTES = 5 * 1024 * 1024          # ~5 MB per foto
MAX_BODY_BYTES = 12 * 1024 * 1024         # batas total satu request multipart
MAX_BACKUP_BYTES = 200 * 1024 * 1024      # berkas backup boleh besar (banyak foto)
FOTO_NAME_RE = re.compile(r"[a-f0-9]{24}\.(?:jpg|jpeg|png|webp)")
KUNCI_EXCEL = "pemilu-osis-terkunci"  # kata kunci proteksi sheet Excel bukti

NAMA_BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
              "Agustus", "September", "Oktober", "November", "Desember"]

# ---------------------------------------------------------------------------
# Waktu — semua keputusan "sekarang" lewat satu helper ini (Asia/Jakarta).
# Untuk pengujian, env var VOTE_NOW_OVERRIDE (format ISO, mis.
# '2026-11-10T10:00:00') menggantikan jam asli; env dibaca SETIAP
# pemanggilan, jadi server uji cukup di-restart dengan env baru.
# Ini hook khusus tes — jangan dipakai di produksi.
# ---------------------------------------------------------------------------


def waktu_sekarang():
    ov = os.environ.get("VOTE_NOW_OVERRIDE", "").strip()
    if ov:
        try:
            dt = datetime.datetime.fromisoformat(ov)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=TZ_JKT)
            return dt.astimezone(TZ_JKT)
        except ValueError:
            pass
    return datetime.datetime.now(TZ_JKT)


def parse_jadwal(nilai):
    """'2026-11-10T08:00' (dari input datetime-local) -> datetime Jakarta.
    Nilai kosong/rusak -> None."""
    nilai = (nilai or "").strip()
    if not nilai:
        return None
    try:
        dt = datetime.datetime.fromisoformat(nilai)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=TZ_JKT)
    return dt.astimezone(TZ_JKT)


def fmt_jadwal(dt):
    """datetime -> '10 November 2026 pukul 08.00 WIB'."""
    if dt is None:
        return "-"
    return (f"{dt.day} {NAMA_BULAN[dt.month]} {dt.year} "
            f"pukul {dt:%H}.{dt:%M} WIB")


def fmt_persen(nilai):
    """71.5 -> '71,5%' (koma desimal gaya Indonesia)."""
    return f"{nilai:.1f}".replace(".", ",") + "%"


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    salt TEXT NOT NULL,
    nama TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS kandidat (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nomor_urut INTEGER UNIQUE NOT NULL,     -- nomor urut surat suara (1, 2, 3, ...)
    nama TEXT NOT NULL,
    kelas TEXT NOT NULL DEFAULT '',
    visi TEXT NOT NULL DEFAULT '',          -- visi singkat satu-dua baris
    foto TEXT                               -- nama berkas foto di folder uploads
);
CREATE TABLE IF NOT EXISTS pemilih (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nis TEXT UNIQUE NOT NULL,               -- Nomor Induk Siswa = identitas pemilih
    nama TEXT NOT NULL,
    kelas TEXT NOT NULL DEFAULT '',
    sudah_memilih INTEGER NOT NULL DEFAULT 0,
    waktu_memilih TEXT                      -- ISO Jakarta, terisi saat suara terkirim
);
CREATE TABLE IF NOT EXISTS suara (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kandidat_id INTEGER NOT NULL REFERENCES kandidat(id),
    waktu TEXT NOT NULL
    -- SENGAJA tanpa kolom pemilih: kerahasiaan suara. Dari tabel ini
    -- TIDAK MUNGKIN diketahui siapa memilih siapa.
);
"""

SETTINGS_AWAL = {
    "nama_pemilu": "Pemilihan Ketua OSIS",
    "nama_sekolah": "",
    "logo": "",
    "waktu_buka": "",
    "waktu_tutup": "",
    "hasil_live": "1",
}


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                             bytes.fromhex(salt), 200_000)
    return dk.hex(), salt


def verify_password(password, password_hash, salt):
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                             bytes.fromhex(salt), 200_000)
    return secrets.compare_digest(dk.hex(), password_hash)


def get_setting(conn, key, default=""):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(conn, key, value):
    conn.execute("INSERT INTO settings (key, value) VALUES (?,?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                 (key, value))


def ambil_pengaturan(conn):
    return {k: get_setting(conn, k, v) for k, v in SETTINGS_AWAL.items()}


def status_voting(peng, now=None):
    """Status jendela voting: (status, buka, tutup).
    status: 'belum-diatur' | 'belum-dibuka' | 'berlangsung' | 'selesai'.
    Voting hanya terbuka bila KEDUA jadwal terisi dan buka <= now <= tutup."""
    now = now or waktu_sekarang()
    buka = parse_jadwal(peng.get("waktu_buka", ""))
    tutup = parse_jadwal(peng.get("waktu_tutup", ""))
    if buka is None or tutup is None:
        return "belum-diatur", buka, tutup
    if now < buka:
        return "belum-dibuka", buka, tutup
    if now > tutup:
        return "selesai", buka, tutup
    return "berlangsung", buka, tutup


def hasil_terlihat(peng, status):
    """Apakah rincian suara per kandidat boleh tampil ke publik (/layar)?
    Boleh bila toggle hasil_live menyala, ATAU voting sudah selesai."""
    return (peng.get("hasil_live", "1") == "1") or status == "selesai"


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    conn = db()
    conn.executescript(SCHEMA)
    conn.commit()
    row = conn.execute("SELECT COUNT(*) c FROM users").fetchone()
    if row["c"] == 0:
        # Akun admin AWAL khusus pengujian lokal. WAJIB diganti sebelum produksi.
        ph, salt = hash_password("admin123")
        conn.execute(
            "INSERT INTO users (username, password_hash, salt, nama) VALUES (?,?,?,?)",
            ("admin", ph, salt, "Administrator"))
        conn.commit()
        print("[init] Akun admin awal dibuat: username=admin (password ada di README). "
              "WAJIB ganti password sebelum dipakai beneran.")
    for k, v in SETTINGS_AWAL.items():
        conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)", (k, v))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Session (server-side, cookie hanya menyimpan token acak)
# Dua tipe sesi yang TERPISAH JELAS:
#   {"tipe": "admin",   "user_id": ...}    -> lolos penjaga admin
#   {"tipe": "pemilih", "pemilih_id": ...} -> hanya untuk alur memilih;
#                                             TIDAK PERNAH lolos penjaga admin
#   {"tipe": "selesai"}                    -> penanda sekali pakai halaman sukses
# ---------------------------------------------------------------------------
LOGIN_SESSIONS = {}  # token -> dict sesi


def esc(v):
    return html.escape(str(v if v is not None else ""), quote=True)


def slug_nama(teks):
    s = re.sub(r"[^a-z0-9]+", "-", (teks or "").lower()).strip("-")
    return s or "pemilu-osis"


# ---------------------------------------------------------------------------
# Unggah foto (multipart, standard library saja)
# ---------------------------------------------------------------------------
def save_upload(files, field):
    """Simpan satu berkas foto dari form multipart.
    Kembalikan (nama_berkas, pesan_error). nama_berkas None bila tidak ada
    berkas yang dipilih (itu bukan error). Nama berkas hasil adalah token
    acak + ekstensi, jadi tidak bisa ditebak/didaftar dari luar."""
    if field not in files:
        return None, ""
    fname, data = files[field]
    if not fname:
        return None, ""
    ext = os.path.splitext(fname)[1].lower()
    if ext not in ALLOWED_EXT:
        return None, "Format foto harus JPG, JPEG, PNG, atau WebP."
    if not data:
        return None, "Berkas foto kosong."
    if len(data) > MAX_FOTO_BYTES:
        return None, "Ukuran foto melebihi 5 MB."
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    new_name = secrets.token_hex(12) + ext
    with open(os.path.join(UPLOADS_DIR, new_name), "wb") as fh:
        fh.write(data)
    return new_name, ""


def delete_upload(name):
    """Hapus berkas unggahan dari folder uploads (hanya nama berpola token)."""
    if name and FOTO_NAME_RE.fullmatch(name):
        try:
            os.remove(os.path.join(UPLOADS_DIR, name))
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Merek & ikon
# ---------------------------------------------------------------------------
def logo_svg(kelas="logo-mark"):
    """Tanda bawaan: kotak suara (ballot box) dengan kertas bercoblos."""
    return (f'<span class="{kelas}" aria-hidden="true">'
            '<svg viewBox="0 0 24 24" width="26" height="26" fill="none" '
            'stroke="currentColor" stroke-width="1.8" stroke-linecap="round" '
            'stroke-linejoin="round">'
            '<path d="M9.5 2.8h5V9h-5z" fill="currentColor" stroke="none" opacity=".9"/>'
            '<path d="M10.8 5.6l1 1 1.8-2"/>'
            '<path d="M4 12.5h16l-1.7 8.7a1.5 1.5 0 0 1-1.5 1.2H7.2a1.5 1.5 0 0 1-1.5-1.2z"/>'
            '<path d="M7.5 12.5a4.5 2.4 0 0 1 9 0"/></svg></span>')


def logo_html(logo, css_class="logo-img"):
    """Logo sekolah dari Pengaturan; bila belum diatur, tanda kotak suara."""
    if logo:
        return f'<img class="{css_class}" src="/foto/{esc(logo)}" alt="Logo sekolah">'
    return logo_svg("logo-mark" + (" kecil" if "kecil" in css_class else "")
                    + (" besar" if "besar" in css_class else ""))


_IKON = {
    "dashboard": '<path d="M4 4h7v7H4z"/><path d="M13 4h7v4h-7z"/><path d="M13 10h7v10h-7z"/><path d="M4 13h7v7H4z"/>',
    "kandidat": '<circle cx="12" cy="8" r="3.6"/><path d="M5 20c.8-3.6 3.6-5.4 7-5.4s6.2 1.8 7 5.4"/>',
    "pemilih": ('<circle cx="9" cy="8.5" r="3.2"/><path d="M3.5 19.5c.6-3 2.8-4.7 5.5-4.7s4.9 1.7 5.5 4.7"/>'
                '<circle cx="17" cy="9.5" r="2.4"/><path d="M16 14.9c2.1.3 3.8 1.8 4.4 4.1"/>'),
    "hasil": '<path d="M4 20V10"/><path d="M10 20V4"/><path d="M16 20v-8"/><path d="M22 20H2"/>',
    "layar": '<rect x="3" y="4.5" width="18" height="12" rx="1.6"/><path d="M9.5 20.5h5M12 16.5v4"/>',
    "atur": ('<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/>'
             '<circle cx="15" cy="7" r="2.2"/><circle cx="9" cy="17" r="2.2"/>'),
    "backup": '<path d="M12 3v11"/><path d="M7 10l5 5 5-5"/><path d="M4 20h16"/>',
    "kunci": '<circle cx="8" cy="15" r="4.5"/><path d="M11.5 11.5 20 3"/><path d="M16 7l3 3"/>',
    "keluar": '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/>',
    "trofi": '<path d="M7 4h10v5a5 5 0 0 1-10 0z"/><path d="M7 5H4a3 3 0 0 0 3 4M17 5h3a3 3 0 0 1-3 4"/><path d="M12 14v3M8.5 20h7M12 17l-1 3h2z"/>',
    "cek": '<path d="M4 12.5l5 5L20 6.5"/>',
    "jam": '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
}


def ikon(nama):
    return (f'<svg class="ikon" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            f'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
            f'{_IKON.get(nama, "")}</svg>')


def _merek():
    """(pengaturan) dari database untuk kebutuhan layout."""
    conn = db()
    peng = ambil_pengaturan(conn)
    conn.close()
    return peng


# ---------------------------------------------------------------------------
# Agregasi hasil — dipakai dashboard admin, halaman hasil, dan /layar.
# Semua angka berasal dari tabel suara (agregat) + tabel pemilih
# (partisipasi). TIDAK ADA kueri yang menghubungkan suara ke pemilih.
# ---------------------------------------------------------------------------
def agregat():
    conn = db()
    peng = ambil_pengaturan(conn)
    kand = conn.execute(
        """SELECT k.*, COALESCE(s.jml, 0) AS suara
           FROM kandidat k
           LEFT JOIN (SELECT kandidat_id, COUNT(*) jml FROM suara
                      GROUP BY kandidat_id) s ON s.kandidat_id = k.id
           ORDER BY suara DESC, k.nomor_urut ASC""").fetchall()
    total_pemilih = conn.execute("SELECT COUNT(*) c FROM pemilih").fetchone()["c"]
    sudah = conn.execute(
        "SELECT COUNT(*) c FROM pemilih WHERE sudah_memilih=1").fetchone()["c"]
    total_suara = conn.execute("SELECT COUNT(*) c FROM suara").fetchone()["c"]
    per_kelas = conn.execute(
        """SELECT kelas, COUNT(*) total,
                  SUM(CASE WHEN sudah_memilih=1 THEN 1 ELSE 0 END) sudah
           FROM pemilih GROUP BY kelas ORDER BY kelas""").fetchall()
    conn.close()
    partisipasi = (total_suara / total_pemilih * 100) if total_pemilih else 0.0
    return {
        "peng": peng, "kandidat": kand, "total_pemilih": total_pemilih,
        "sudah": sudah, "belum": total_pemilih - sudah,
        "total_suara": total_suara, "partisipasi": partisipasi,
        "per_kelas": per_kelas,
    }


def podium_html(kand, tampil_suara=True):
    """Podium 3 teratas: peringkat 1 (emas) di tengah, 2 kiri, 3 kanan."""
    top = list(kand[:3])
    if not top:
        return '<p class="redup">Belum ada kandidat.</p>'

    def kartu(k, peringkat):
        foto = (f'<img src="/foto/{esc(k["foto"])}" alt="Foto {esc(k["nama"])}">'
                if k["foto"] else
                f'<span class="podium-inisial">{esc((k["nama"] or "?")[:1].upper())}</span>')
        suara = (f'<div class="podium-suara">{k["suara"]} suara &bull; '
                 f'{fmt_persen(k["suara"] / max(1, sum(x["suara"] for x in kand)) * 100)}</div>'
                 if tampil_suara else "")
        unggul = ""
        if peringkat == 1 and tampil_suara and len(kand) > 1:
            selisih = k["suara"] - kand[1]["suara"]
            if selisih > 0:
                unggul = f'<div class="podium-unggul">Unggul +{selisih} suara</div>'
        return (f'<div class="podium-kartu p{peringkat}">'
                f'<div class="podium-foto">{foto}<span class="podium-rank">{peringkat}</span></div>'
                f'<div class="podium-nama">{esc(k["nama"])}</div>'
                f'<div class="podium-meta">No. {k["nomor_urut"]:02d} &bull; {esc(k["kelas"])}</div>'
                f'{suara}{unggul}</div>')

    urutan = []
    if len(top) >= 2:
        urutan.append(kartu(top[1], 2))
    urutan.append(kartu(top[0], 1))
    if len(top) >= 3:
        urutan.append(kartu(top[2], 3))
    return f'<div class="podium">{"".join(urutan)}</div>'


def bar_kandidat_html(kand, tampil_suara=True):
    total = max(1, sum(k["suara"] for k in kand))
    baris = []
    for k in kand:
        p = k["suara"] / total * 100 if tampil_suara else 0
        angka = (f'<span class="bar-angka">{k["suara"]} suara ({fmt_persen(p)})</span>'
                 if tampil_suara else "")
        baris.append(
            f'<div class="bar-baris"><span class="bar-nama">'
            f'No. {k["nomor_urut"]:02d} &bull; {esc(k["nama"])}</span>'
            f'<span class="bar-track"><span class="bar-isi" style="width:{p:.1f}%"></span></span>'
            f'{angka}</div>')
    return "".join(baris) or '<p class="redup">Belum ada kandidat.</p>'


def donat_svg(persen):
    """Diagram donat partisipasi (SVG, stroke-dasharray)."""
    import math
    keliling = 2 * math.pi * 54
    terisi = max(0.0, min(100.0, persen)) / 100 * keliling
    return (f'<svg class="donat" viewBox="0 0 140 140" role="img" '
            f'aria-label="Partisipasi {fmt_persen(persen)}">'
            f'<circle cx="70" cy="70" r="54" class="donat-latar"/>'
            f'<circle cx="70" cy="70" r="54" class="donat-isi" '
            f'stroke-dasharray="{terisi:.1f} {keliling:.1f}"/>'
            f'<text x="70" y="68" class="donat-angka">{fmt_persen(persen)}</text>'
            f'<text x="70" y="88" class="donat-label">sudah memilih</text></svg>')


def sisa_waktu_html(status, buka, tutup, now):
    """Kartu hitungan mundur: menuju buka / menuju tutup / selesai.
    Server merender teks awal; app.js memperbarui tiap detik dari
    data-target (epoch milidetik)."""
    if status == "belum-diatur":
        return ('<div class="stat-angka kecil">Belum diatur</div>'
                '<div class="stat-sub">Atur jadwal di Pengaturan</div>')
    if status == "selesai":
        return ('<div class="stat-angka kecil">Selesai</div>'
                f'<div class="stat-sub">Berakhir {fmt_jadwal(tutup)}</div>')
    target = buka if status == "belum-dibuka" else tutup
    label = "Mulai dalam" if status == "belum-dibuka" else "Sisa Waktu Voting"
    detik = max(0, int((target - now).total_seconds()))
    hari, sisa = divmod(detik, 86400)
    jam, sisa = divmod(sisa, 3600)
    menit, dtk = divmod(sisa, 60)
    teks = (f"{hari} hari " if hari else "") + f"{jam:02d}:{menit:02d}:{dtk:02d}"
    epoch_ms = int(target.timestamp() * 1000)
    return (f'<div class="stat-label">{label}</div>'
            f'<div class="stat-angka hitung" data-target="{epoch_ms}">{teks}</div>'
            f'<div class="stat-sub">{"Dibuka" if status == "belum-dibuka" else "Berakhir"} '
            f'{fmt_jadwal(target)}</div>')


# ---------------------------------------------------------------------------
# Layout ADMIN — tema GELAP (navy + toska + emas), sidebar desktop +
# bilah atas & navigasi bawah di HP. Dipakai juga halaman masuk (polos).
# ---------------------------------------------------------------------------
def layout_admin(user, title, body, flash="", aktif="", peringatan=""):
    peng = _merek()
    nama_pemilu = peng["nama_pemilu"] or "Pemilihan Ketua OSIS"
    logo = peng["logo"]
    flash_html = f'<div class="flash">{esc(flash)}</div>' if flash else ""
    warn_html = (f'<div class="flash warn">{peringatan}</div>' if peringatan else "")
    if not user:
        return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} - {esc(nama_pemilu)}</title>
<link rel="stylesheet" href="/static/style.css"></head>
<body class="polos gelap"><main class="polos-main">
<div class="brand-center">{logo_html(logo, "logo-img besar")}
<div class="brand-nama-besar">{esc(nama_pemilu)}</div>
<div class="brand-sub">e-Voting Sekolah &bull; Panel Admin</div></div>
{flash_html}{body}</main>
<script src="/static/app.js"></script>
</body></html>"""
    item = [
        ("dashboard", "/admin", "Dashboard", "dashboard"),
        ("kandidat", "/admin/kandidat", "Kandidat", "kandidat"),
        ("pemilih", "/admin/pemilih", "Pemilih", "pemilih"),
        ("hasil", "/admin/hasil", "Hasil Suara", "hasil"),
        ("layar", "/layar", "Layar Tampil", "layar"),
        ("pengaturan", "/admin/pengaturan", "Pengaturan", "atur"),
        ("backup", "/admin/backup", "Backup & Pulihkan", "backup"),
        ("password", "/ganti-password", "Ganti Password", "kunci"),
        ("keluar", "/keluar", "Keluar", "keluar"),
    ]

    def tautan(it):
        kunci, href, label, nama_ikon = it
        cls = ' class="aktif"' if kunci == aktif else ""
        target = ' target="_blank" rel="noopener"' if kunci == "layar" else ""
        return f'<a href="{href}"{cls}{target}>{ikon(nama_ikon)}<span>{esc(label)}</span></a>'

    samping = "".join(tautan(it) for it in item)
    bawah_item = [item[0], item[1], item[2], item[3], item[8]]
    bawah = "".join(tautan(it) for it in bawah_item)
    awal = (user["nama"] or "A").strip()[:1].upper()
    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} - {esc(nama_pemilu)}</title>
<link rel="stylesheet" href="/static/style.css"></head>
<body class="gelap"><div class="shell">
  <aside class="sidebar">
    <div class="brand">{logo_html(logo)}<div class="brand-teks"><b>PEMILU OSIS</b><small>e-Voting Sekolah</small></div></div>
    <nav class="side-nav">{samping}</nav>
    <div class="side-user"><span class="avatar">{esc(awal)}</span><div><b>{esc(user['nama'])}</b><small>Administrator</small></div></div>
  </aside>
  <div class="shell-main">
    <div class="topbar">{logo_html(logo, "logo-img kecil")}<b>PEMILU OSIS</b><span class="top-user">{esc(user['nama'])}</span></div>
    <main>{flash_html}{warn_html}{body}</main>
    <footer class="kaki"><small>PEMILU OSIS &bull; e-Voting Sekolah &bull; Suara bersifat rahasia</small></footer>
  </div>
</div>
<nav class="bottom-nav">{bawah}</nav>
<script src="/static/app.js"></script>
</body></html>"""


# ---------------------------------------------------------------------------
# Layout PEMILIH — tema TERANG (biru + putih + tombol hijau), satu kolom
# ramah HP. TANPA menu admin sama sekali.
# ---------------------------------------------------------------------------
def layout_pemilih(title, body, flash="", pemilih=None):
    peng = _merek()
    nama_pemilu = peng["nama_pemilu"] or "Pemilihan Ketua OSIS"
    nama_sekolah = peng["nama_sekolah"]
    logo = peng["logo"]
    flash_html = f'<div class="flash terang">{esc(flash)}</div>' if flash else ""
    sapaan = ""
    if pemilih:
        sapaan = (f'<div class="pemilih-chip">Terhubung sebagai '
                  f'<b>{esc(pemilih["nama"])}</b> &bull; NIS {esc(pemilih["nis"])} '
                  f'&bull; {esc(pemilih["kelas"])}</div>')
    sub = esc(nama_sekolah) if nama_sekolah else "e-Voting Sekolah"
    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} - {esc(nama_pemilu)}</title>
<link rel="stylesheet" href="/static/style.css"></head>
<body class="terang">
<header class="voter-head">
  <div class="voter-head-dalam">{logo_html(logo)}
  <div><h1>{esc(nama_pemilu)}</h1><p>{sub}</p></div></div>
</header>
<main class="voter-main">{flash_html}{sapaan}{body}</main>
<footer class="voter-kaki"><small>Aman &bull; Rahasia &bull; Satu NIS satu suara</small></footer>
<script src="/static/app.js"></script>
</body></html>"""


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "PemiluOSIS/1.0"

    def log_message(self, fmt, *args):  # lebih senyap
        pass

    # -- util ---------------------------------------------------------------
    def send_html(self, content, status=200, extra_headers=None):
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra_headers or []):
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def send_bytes(self, data, ctype, filename=None):
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, path, flash="", set_cookie=None):
        headers = []
        if set_cookie:
            headers.append(("Set-Cookie", set_cookie))
        if flash:
            path += ("&" if "?" in path else "?") + "msg=" + urllib.parse.quote(flash)
        headers.append(("Location", path))
        self.send_html("", status=303, extra_headers=headers)

    # -- sesi ---------------------------------------------------------------
    def _token(self):
        cookie = SimpleCookie(self.headers.get("Cookie", ""))
        morsel = cookie.get("vote_session")
        return morsel.value if morsel else None

    def _sesi(self):
        token = self._token()
        return LOGIN_SESSIONS.get(token) if token else None

    def _buat_sesi(self, data):
        token = secrets.token_urlsafe(32)
        LOGIN_SESSIONS[token] = data
        return f"vote_session={token}; HttpOnly; Path=/; SameSite=Lax"

    def _hapus_sesi(self):
        token = self._token()
        if token:
            LOGIN_SESSIONS.pop(token, None)

    def current_user(self):
        """Baris users bila sesi ini sesi ADMIN; selain itu None.
        Sesi pemilih TIDAK PERNAH dianggap admin."""
        sesi = self._sesi()
        if not sesi or sesi.get("tipe") != "admin":
            return None
        conn = db()
        row = conn.execute("SELECT * FROM users WHERE id=?",
                           (sesi.get("user_id"),)).fetchone()
        conn.close()
        return row

    def current_pemilih(self):
        """Baris pemilih bila sesi ini sesi PEMILIH; selain itu None."""
        sesi = self._sesi()
        if not sesi or sesi.get("tipe") != "pemilih":
            return None
        conn = db()
        row = conn.execute("SELECT * FROM pemilih WHERE id=?",
                           (sesi.get("pemilih_id"),)).fetchone()
        conn.close()
        return row

    def peringatan_password(self, user):
        """Banner peringatan selama password admin masih password awal."""
        if user and verify_password("admin123", user["password_hash"], user["salt"]):
            return ("PERINGATAN: akun admin masih memakai password awal. "
                    "Segera ganti lewat menu Ganti Password sebelum aplikasi "
                    "dipakai beneran.")
        return ""

    # -- form ---------------------------------------------------------------
    def form(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8") if length else ""
        parsed = urllib.parse.parse_qs(raw)
        return {k: v[0] for k, v in parsed.items()}

    def parse_multipart(self, max_bytes=MAX_BODY_BYTES):
        """Parser multipart/form-data minimal (standard library saja).
        Kembalikan (fields, files); files: nama_field -> (nama_berkas, bytes)."""
        ctype = self.headers.get("Content-Type", "")
        m = re.search(r"boundary=(.+)", ctype)
        if not m:
            return {}, {}
        boundary = m.group(1).strip().strip('"')
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length > max_bytes:
            raise ValueError("Ukuran unggahan melebihi batas yang diizinkan.")
        body = self.rfile.read(length) if length else b""
        delim = b"--" + boundary.encode("utf-8")
        fields, files = {}, {}
        for part in body.split(delim):
            if part.startswith(b"\r\n"):
                part = part[2:]
            if part.endswith(b"--"):
                part = part[:-2]
            if part.endswith(b"\r\n"):
                part = part[:-2]
            if not part or b"\r\n\r\n" not in part:
                continue
            head, data = part.split(b"\r\n\r\n", 1)
            if data.endswith(b"\r\n"):
                data = data[:-2]
            headers = head.decode("utf-8", "replace")
            name_m = re.search(r'name="([^"]+)"', headers)
            if not name_m:
                continue
            name = name_m.group(1)
            fn_m = re.search(r'filename="([^"]*)"', headers)
            if fn_m is not None:
                files[name] = (fn_m.group(1), data)
            else:
                fields[name] = data.decode("utf-8", "replace")
        return fields, files

    def read_form(self):
        """Baca POST body: multipart -> (fields, files); urlencoded -> (fields, {})."""
        ctype = self.headers.get("Content-Type", "")
        if ctype.startswith("multipart/form-data"):
            return self.parse_multipart()
        return self.form(), {}

    # -- routing ------------------------------------------------------------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path, qs = parsed.path, urllib.parse.parse_qs(parsed.query)
        flash = (qs.get("msg") or [""])[0]
        try:
            self.route_get(path, qs, flash)
        except Exception as e:  # jangan bocorkan stack ke user, tapi catat di log server
            import traceback; traceback.print_exc()
            self.send_html(f"<p>Terjadi kesalahan: {esc(e)}</p>", status=500)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        try:
            self.route_post(parsed.path)
        except Exception as e:
            import traceback; traceback.print_exc()
            self.send_html(f"<p>Terjadi kesalahan: {esc(e)}</p>", status=500)

    def route_get(self, path, qs, flash):
        if path in ("/static/style.css", "/static/app.js"):
            fname = os.path.join(BASE_DIR, path.lstrip("/"))
            ctype = "text/css" if path.endswith(".css") else "application/javascript"
            with open(fname, "rb") as fh:
                data = fh.read()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path.startswith("/foto/"):
            return self.serve_foto(path[len("/foto/"):])
        if path == "/":
            return self.voter_home(flash)
        if path == "/surat-suara":
            return self.surat_suara(flash)
        if path == "/sukses":
            return self.sukses(flash)
        if path == "/layar":
            return self.layar(flash)
        if path == "/login":
            body = """<section class="panel kartu-masuk"><h2>Masuk Admin</h2>
            <form method="post" action="/login">
              <label>Username <input name="username" required autocomplete="username"></label>
              <label>Password <input name="password" type="password" required autocomplete="current-password"></label>
              <button type="submit" class="btn-emas">Masuk</button></form>
            <p><small>Akun uji lokal ada di README proyek. Jangan pakai akun awal di produksi.</small></p>
            <p><a href="/">Ke halaman pemilih</a></p></section>"""
            return self.send_html(layout_admin(None, "Masuk Admin", body, flash))
        if path in ("/logout", "/keluar"):
            self._hapus_sesi()
            return self.redirect("/", "Sampai jumpa.",
                                 set_cookie="vote_session=; Max-Age=0; Path=/")

        user = self.current_user()
        if not user:
            return self.redirect("/login", "Silakan masuk dulu.")
        peringatan = self.peringatan_password(user)
        if path == "/admin":
            return self.admin_dashboard(user, flash, peringatan)
        if path == "/admin/kandidat":
            return self.admin_kandidat_list(user, qs, flash, peringatan)
        if path == "/admin/kandidat/baru":
            return self.admin_kandidat_form(user, None, flash, peringatan)
        if path == "/admin/pemilih":
            return self.admin_pemilih_list(user, qs, flash, peringatan)
        if path == "/admin/pemilih/baru":
            return self.admin_pemilih_form(user, None, flash, peringatan)
        if path == "/admin/hasil":
            return self.admin_hasil(user, flash, peringatan)
        if path == "/admin/hasil/hasil.csv":
            return self.csv_hasil(user)
        if path == "/admin/hasil/partisipasi.csv":
            return self.csv_partisipasi(user)
        if path == "/admin/hasil/hasil.xlsx":
            return self.excel_hasil(user)
        if path == "/admin/berita-acara":
            return self.berita_acara(user)
        if path == "/admin/pengaturan":
            return self.admin_pengaturan(user, flash, peringatan)
        if path == "/admin/backup":
            return self.admin_backup(user, flash, peringatan)
        if path == "/admin/backup/unduh":
            return self.admin_backup_unduh(user)
        if path == "/ganti-password":
            return self.ganti_password_form(user, flash, peringatan)
        if path.startswith("/admin/kandidat/"):
            try:
                kid = int(path.split("/")[3])
            except (IndexError, ValueError):
                return self.not_found(user)
            return self.admin_kandidat_form(user, kid, flash, peringatan)
        if path.startswith("/admin/pemilih/"):
            try:
                pid = int(path.split("/")[3])
            except (IndexError, ValueError):
                return self.not_found(user)
            return self.admin_pemilih_form(user, pid, flash, peringatan)
        return self.not_found(user)

    def not_found(self, user):
        return self.send_html(layout_admin(user, "404",
                                           "<h2>404</h2><p>Halaman tidak ditemukan.</p>"),
                              status=404)

    def route_post(self, path):
        if path == "/login":
            return self.post_login()
        if path == "/cek-nis":
            return self.post_cek_nis()
        if path == "/konfirmasi":
            return self.post_konfirmasi()
        if path == "/suara":
            return self.post_suara()
        user = self.current_user()
        if not user:
            return self.redirect("/login", "Silakan masuk dulu.")
        if path == "/admin/backup/pulihkan":
            # Berkas backup bisa jauh lebih besar dari unggahan foto biasa,
            # jadi dibaca dengan batas ukuran khusus backup.
            try:
                _f, files = self.parse_multipart(MAX_BACKUP_BYTES)
            except ValueError as e:
                return self.redirect("/admin/backup", str(e))
            return self.post_backup_pulihkan(user, files)
        f, files = self.read_form()
        if path == "/ganti-password":
            return self.post_ganti_password(user, f)
        if path == "/admin/pengaturan":
            return self.post_pengaturan(user, f)
        if path == "/admin/pengaturan/logo":
            return self.post_logo(user, files)
        if path == "/admin/pengaturan/logo-hapus":
            return self.post_logo_hapus(user)
        if path == "/admin/kandidat/baru":
            return self.post_kandidat_simpan(user, None, f, files)
        if path == "/admin/pemilih/baru":
            return self.post_pemilih_simpan(user, None, f)
        if path == "/admin/pemilih/impor":
            return self.post_pemilih_impor(user, f)
        if path.startswith("/admin/kandidat/"):
            parts = path.split("/")
            try:
                kid = int(parts[3])
            except (IndexError, ValueError):
                return self.not_found(user)
            aksi = parts[4] if len(parts) > 4 else ""
            if aksi == "edit":
                return self.post_kandidat_simpan(user, kid, f, files)
            if aksi == "hapus":
                return self.post_kandidat_hapus(user, kid)
            if aksi == "foto-hapus":
                return self.post_kandidat_foto_hapus(user, kid)
        if path.startswith("/admin/pemilih/"):
            parts = path.split("/")
            try:
                pid = int(parts[3])
            except (IndexError, ValueError):
                return self.not_found(user)
            aksi = parts[4] if len(parts) > 4 else ""
            if aksi == "edit":
                return self.post_pemilih_simpan(user, pid, f)
            if aksi == "hapus":
                return self.post_pemilih_hapus(user, pid)
        return self.not_found(user)

    # -- berkas foto ----------------------------------------------------------
    def serve_foto(self, raw_name):
        """Foto hanya disajikan lewat route ini, dengan nama berkas token acak.
        DAFTAR PUTIH: nama juga harus tercatat sebagai foto kandidat atau
        logo situs di database."""
        name = urllib.parse.unquote(raw_name)
        if not FOTO_NAME_RE.fullmatch(name):
            return self.send_html("Tidak ditemukan.", status=404)
        conn = db()
        tercatat = conn.execute(
            """SELECT 1 WHERE EXISTS(SELECT 1 FROM kandidat WHERE foto=?)
               OR EXISTS(SELECT 1 FROM settings WHERE key='logo' AND value=?)""",
            (name, name)).fetchone()
        conn.close()
        if not tercatat:
            return self.send_html("Tidak ditemukan.", status=404)
        fpath = os.path.join(UPLOADS_DIR, name)
        if not os.path.isfile(fpath):
            return self.send_html("Tidak ditemukan.", status=404)
        ext = os.path.splitext(name)[1].lower()
        ctype = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                 ".png": "image/png", ".webp": "image/webp"}[ext]
        with open(fpath, "rb") as fh:
            data = fh.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "private, max-age=3600")
        self.end_headers()
        self.wfile.write(data)

    # -- auth admin -----------------------------------------------------------
    def post_login(self):
        f = self.form()
        conn = db()
        row = conn.execute("SELECT * FROM users WHERE username=?",
                           (f.get("username", "").strip(),)).fetchone()
        conn.close()
        if row and verify_password(f.get("password", ""), row["password_hash"], row["salt"]):
            cookie = self._buat_sesi({"tipe": "admin", "user_id": row["id"]})
            return self.redirect("/admin", f"Selamat datang, {row['nama']}.",
                                 set_cookie=cookie)
        return self.redirect("/login", "Username atau password salah.")

    def ganti_password_form(self, user, flash, peringatan=""):
        body = """<h2>Ganti Password</h2>
        <section class="panel"><form method="post" action="/ganti-password">
          <label>Password lama <input name="lama" type="password" required autocomplete="current-password"></label>
          <label>Password baru (min. 8 karakter) <input name="baru" type="password" required minlength="8" autocomplete="new-password"></label>
          <label>Ulangi password baru <input name="ulang" type="password" required minlength="8" autocomplete="new-password"></label>
          <button type="submit" class="btn-emas">Simpan Password Baru</button></form></section>"""
        return self.send_html(layout_admin(user, "Ganti Password", body, flash,
                                           aktif="password", peringatan=peringatan))

    def post_ganti_password(self, user, f):
        if not verify_password(f.get("lama", ""), user["password_hash"], user["salt"]):
            return self.redirect("/ganti-password", "Password lama salah.")
        baru = f.get("baru", "")
        if len(baru) < 8:
            return self.redirect("/ganti-password", "Password baru minimal 8 karakter.")
        if baru != f.get("ulang", ""):
            return self.redirect("/ganti-password", "Pengulangan password baru tidak sama.")
        ph, salt = hash_password(baru)
        conn = db()
        conn.execute("UPDATE users SET password_hash=?, salt=? WHERE id=?",
                     (ph, salt, user["id"]))
        conn.commit()
        conn.close()
        return self.redirect("/admin", "Password berhasil diganti.")


    # -----------------------------------------------------------------------
    # PEMILIH: halaman depan, cek NIS, surat suara, konfirmasi, kirim suara
    # -----------------------------------------------------------------------
    def _info_jadwal(self, peng, status, buka, tutup):
        if status == "belum-diatur":
            return ('<div class="info-box">Jadwal pemungutan suara <b>belum diatur</b> '
                    'oleh panitia. Silakan kembali lagi nanti.</div>')
        if status == "belum-dibuka":
            return (f'<div class="info-box">Pemungutan suara <b>belum dibuka</b>. '
                    f'Dibuka {fmt_jadwal(buka)}, ditutup {fmt_jadwal(tutup)}.</div>')
        if status == "selesai":
            return (f'<div class="info-box">Pemungutan suara <b>sudah ditutup</b> '
                    f'({fmt_jadwal(tutup)}). Terima kasih atas partisipasinya.</div>')
        return (f'<div class="info-box buka">Pemungutan suara <b>sedang berlangsung</b> '
                f'sampai {fmt_jadwal(tutup)}. Silakan gunakan hak pilihmu.</div>')

    def _kartu_pratinjau(self, kand):
        if not kand:
            return '<p class="redup-terang">Belum ada kandidat yang terdaftar.</p>'
        kartu = []
        for k in kand:
            foto = (f'<img class="kand-foto" src="/foto/{esc(k["foto"])}" '
                    f'alt="Foto {esc(k["nama"])}">' if k["foto"] else
                    f'<span class="kand-foto inisial">{esc((k["nama"] or "?")[:1].upper())}</span>')
            kartu.append(f"""<article class="kand-kartu">
              {foto}
              <div class="kand-isi"><span class="no-badge">{k['nomor_urut']:02d}</span>
              <h3>{esc(k['nama'])}</h3>
              <p class="kand-kelas">{esc(k['kelas'])}</p>
              <p class="kand-visi">{esc(k['visi'])}</p></div></article>""")
        return "".join(kartu)

    def voter_home(self, flash):
        conn = db()
        peng = ambil_pengaturan(conn)
        kand = conn.execute(
            "SELECT * FROM kandidat ORDER BY nomor_urut").fetchall()
        conn.close()
        now = waktu_sekarang()
        status, buka, tutup = status_voting(peng, now)
        pemilih = self.current_pemilih()
        lanjut = ""
        if pemilih and not pemilih["sudah_memilih"] and status == "berlangsung":
            lanjut = ('<p><a class="btn-hijau blok" href="/surat-suara">'
                      'Lanjut Memilih &rarr;</a></p>')
        body = f"""{self._info_jadwal(peng, status, buka, tutup)}
        <section class="panel terang"><h2>Cek Identitas Pemilih</h2>
        <p class="redup-terang">Masukkan Nomor Induk Siswa (NIS) untuk
        memverifikasi data pemilih. Satu NIS hanya dapat memilih satu kali.</p>
        <form method="post" action="/cek-nis">
          <label>Nomor Induk Siswa (NIS)
            <input name="nis" inputmode="numeric" autocomplete="off" required
              placeholder="Masukkan NIS Anda"></label>
          <button type="submit" class="btn-hijau blok">Verifikasi NIS</button></form>
        {lanjut}</section>
        <h2 class="judul-seksi">Kandidat Ketua OSIS</h2>
        {self._kartu_pratinjau(kand)}"""
        return self.send_html(layout_pemilih("Beranda Pemilih", body, flash, pemilih))

    def post_cek_nis(self):
        f = self.form()
        nis = (f.get("nis") or "").strip()
        conn = db()
        peng = ambil_pengaturan(conn)
        pem = conn.execute("SELECT * FROM pemilih WHERE nis=?", (nis,)).fetchone()
        conn.close()
        if not pem:
            return self.redirect("/", "NIS tidak ditemukan. Periksa kembali NIS kamu, "
                                      "atau hubungi panitia.")
        if pem["sudah_memilih"]:
            return self.redirect("/", "Kamu sudah menggunakan hak pilihmu. "
                                      "Terima kasih atas partisipasinya.")
        status, buka, tutup = status_voting(peng)
        if status != "berlangsung":
            if status == "belum-diatur":
                pesan = "Pemungutan suara belum diatur jadwalnya oleh panitia."
            elif status == "belum-dibuka":
                pesan = f"Pemungutan suara belum dibuka. Dibuka {fmt_jadwal(buka)}."
            else:
                pesan = f"Pemungutan suara sudah ditutup ({fmt_jadwal(tutup)})."
            return self.redirect("/", pesan)
        cookie = self._buat_sesi({"tipe": "pemilih", "pemilih_id": pem["id"]})
        return self.redirect("/surat-suara",
                             f"Halo, {pem['nama']}. Silakan pilih kandidatmu.",
                             set_cookie=cookie)

    def surat_suara(self, flash):
        pemilih = self.current_pemilih()
        if not pemilih:
            return self.redirect("/", "Verifikasi NIS dulu untuk mulai memilih.")
        conn = db()
        peng = ambil_pengaturan(conn)
        kand = conn.execute(
            "SELECT * FROM kandidat ORDER BY nomor_urut").fetchall()
        conn.close()
        status, buka, tutup = status_voting(peng)
        if pemilih["sudah_memilih"]:
            body = ('<section class="panel terang"><h2>Sudah Memilih</h2>'
                    '<p>Kamu sudah menggunakan hak pilihmu. Suara yang terkirim '
                    'tidak dapat diubah. Terima kasih atas partisipasinya.</p>'
                    '<p><a class="btn-hijau blok" href="/">Kembali</a></p></section>')
            return self.send_html(layout_pemilih("Sudah Memilih", body, flash, pemilih))
        if status != "berlangsung":
            body = (f'<section class="panel terang"><h2>Pemungutan Suara Ditutup</h2>'
                    f'{self._info_jadwal(peng, status, buka, tutup)}'
                    '<p><a class="btn-hijau blok" href="/">Kembali</a></p></section>')
            return self.send_html(layout_pemilih("Ditutup", body, flash, pemilih))
        if not kand:
            body = ('<section class="panel terang"><h2>Belum Ada Kandidat</h2>'
                    '<p>Panitia belum mendaftarkan kandidat. Silakan kembali lagi.</p></section>')
            return self.send_html(layout_pemilih("Surat Suara", body, flash, pemilih))
        pilihan = []
        for k in kand:
            foto = (f'<img class="kand-foto" src="/foto/{esc(k["foto"])}" '
                    f'alt="Foto {esc(k["nama"])}">' if k["foto"] else
                    f'<span class="kand-foto inisial">{esc((k["nama"] or "?")[:1].upper())}</span>')
            pilihan.append(f"""<label class="kand-kartu pilih">
              <input type="radio" name="kandidat" value="{k['id']}" required>
              {foto}
              <span class="kand-isi"><span class="no-badge">{k['nomor_urut']:02d}</span>
              <span class="kand-nama">{esc(k['nama'])}</span>
              <span class="kand-kelas">{esc(k['kelas'])}</span>
              <span class="kand-visi">{esc(k['visi'])}</span></span>
              <span class="pilih-lingkaran" aria-hidden="true"></span></label>""")
        body = f"""<h2 class="judul-seksi">Surat Suara</h2>
        <p class="redup-terang">Pilih <b>satu</b> kandidat, lalu tekan KIRIM SUARA.
        Kamu akan diminta memeriksa pilihanmu sekali lagi sebelum terkirim.</p>
        <form method="post" action="/konfirmasi">
        {"".join(pilihan)}
        <button type="submit" class="btn-hijau blok besar">KIRIM SUARA</button></form>"""
        return self.send_html(layout_pemilih("Surat Suara", body, flash, pemilih))

    def post_konfirmasi(self):
        pemilih = self.current_pemilih()
        if not pemilih:
            return self.redirect("/", "Verifikasi NIS dulu untuk mulai memilih.")
        f = self.form()
        try:
            kid = int(f.get("kandidat") or "")
        except ValueError:
            return self.redirect("/surat-suara", "Pilih satu kandidat dulu.")
        conn = db()
        k = conn.execute("SELECT * FROM kandidat WHERE id=?", (kid,)).fetchone()
        conn.close()
        if not k:
            return self.redirect("/surat-suara", "Kandidat tidak ditemukan.")
        sesi = self._sesi()
        if sesi is not None:
            sesi["pilihan"] = kid
        foto = (f'<img class="kand-foto besar" src="/foto/{esc(k["foto"])}" '
                f'alt="Foto {esc(k["nama"])}">' if k["foto"] else
                f'<span class="kand-foto inisial besar">{esc((k["nama"] or "?")[:1].upper())}</span>')
        body = f"""<section class="panel terang"><h2>Periksa Pilihanmu</h2>
        <div class="kand-kartu ringkas">{foto}
          <div class="kand-isi"><span class="no-badge">{k['nomor_urut']:02d}</span>
          <h3>{esc(k['nama'])}</h3>
          <p class="kand-kelas">{esc(k['kelas'])}</p>
          <p class="kand-visi">{esc(k['visi'])}</p></div></div>
        <div class="info-box">Pastikan pilihanmu sudah tepat.
        <b>Suara tidak dapat diubah setelah dikirim.</b></div>
        <form method="post" action="/suara">
          <button type="submit" class="btn-hijau blok besar">Kirim Suara Saya</button></form>
        <p><a class="btn-garis blok" href="/surat-suara">Ganti Pilihan</a></p></section>"""
        return self.send_html(layout_pemilih("Konfirmasi", body, "", pemilih))

    def post_suara(self):
        """Kirim suara dalam SATU transaksi: periksa ulang jendela voting
        dan status pemilih, sisipkan surat suara (TANPA identitas pemilih),
        tandai pemilih sudah memilih, lalu hancurkan sesi pemilih."""
        pemilih = self.current_pemilih()
        if not pemilih:
            return self.redirect("/", "Sesi memilih sudah berakhir. "
                                      "Verifikasi NIS dulu untuk mulai memilih.")
        sesi = self._sesi()
        kid = sesi.get("pilihan") if sesi else None
        if not kid:
            return self.redirect("/surat-suara", "Pilih satu kandidat dulu.")
        conn = db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            peng = ambil_pengaturan(conn)
            status, _b, _t = status_voting(peng)
            pem = conn.execute("SELECT * FROM pemilih WHERE id=?",
                               (pemilih["id"],)).fetchone()
            kand = conn.execute("SELECT id FROM kandidat WHERE id=?",
                                (kid,)).fetchone()
            if not pem or pem["sudah_memilih"]:
                conn.rollback()
                conn.close()
                return self.redirect("/", "Kamu sudah menggunakan hak pilihmu. "
                                          "Terima kasih atas partisipasinya.")
            if status != "berlangsung" or not kand:
                conn.rollback()
                conn.close()
                return self.redirect("/", "Pemungutan suara sedang tidak berlangsung, "
                                          "suara tidak tercatat.")
            # Surat suara anonim: HANYA kandidat + waktu. Tanpa pemilih.
            conn.execute("INSERT INTO suara (kandidat_id, waktu) VALUES (?,?)",
                         (kid, waktu_sekarang().isoformat()))
            conn.execute(
                "UPDATE pemilih SET sudah_memilih=1, waktu_memilih=? WHERE id=?",
                (waktu_sekarang().isoformat(), pemilih["id"]))
            conn.commit()
        finally:
            conn.close()
        self._hapus_sesi()
        cookie = self._buat_sesi({"tipe": "selesai"})
        return self.redirect("/sukses", set_cookie=cookie)

    def sukses(self, flash):
        sesi = self._sesi()
        if not sesi or sesi.get("tipe") != "selesai":
            return self.redirect("/")
        self._hapus_sesi()
        body = """<section class="panel terang sukses">
        <div class="sukses-cek">&#10003;</div>
        <h2>Terima kasih — suaramu sudah tercatat.</h2>
        <p>Suaramu bersifat rahasia: tidak ada yang bisa mengetahui pilihanmu.
        Kamu boleh menutup halaman ini sekarang.</p>
        <p><a class="btn-hijau blok" href="/">Kembali ke Beranda</a></p></section>"""
        return self.send_html(layout_pemilih("Berhasil", body))

    # -----------------------------------------------------------------------
    # LAYAR TAMPIL — papan hasil publik layar penuh (tanpa login).
    # Hanya angka agregat; TIDAK ADA identitas pemilih di halaman ini.
    # -----------------------------------------------------------------------
    def layar(self, flash=""):
        agg = agregat()
        peng = agg["peng"]
        now = waktu_sekarang()
        status, buka, tutup = status_voting(peng, now)
        tampil = hasil_terlihat(peng, status)
        nama_pemilu = peng["nama_pemilu"] or "Pemilihan Ketua OSIS"
        nama_sekolah = peng["nama_sekolah"]
        if status == "belum-diatur":
            status_teks = "Jadwal belum diatur"
        elif status == "belum-dibuka":
            status_teks = "Voting belum dibuka"
        elif status == "berlangsung":
            status_teks = "LIVE — voting berlangsung"
        else:
            status_teks = "Voting selesai"
        if tampil:
            tengah = (f'<section class="layar-panel podium-panel"><h3>Hasil Live — '
                      f'Podium Kandidat</h3>{podium_html(agg["kandidat"], True)}</section>'
                      f'<section class="layar-panel"><h3>Suara per Kandidat</h3>'
                      f'{bar_kandidat_html(agg["kandidat"], True)}</section>')
        else:
            kartu = []
            for k in agg["kandidat"]:
                foto = (f'<img src="/foto/{esc(k["foto"])}" alt="Foto {esc(k["nama"])}">'
                        if k["foto"] else
                        f'<span class="podium-inisial">{esc((k["nama"] or "?")[:1].upper())}</span>')
                kartu.append(f'<div class="podium-kartu polos">{foto}'
                             f'<div class="podium-nama">{esc(k["nama"])}</div>'
                             f'<div class="podium-meta">No. {k["nomor_urut"]:02d} &bull; '
                             f'{esc(k["kelas"])}</div></div>')
            tengah = (f'<section class="layar-panel podium-panel"><h3>Kandidat</h3>'
                      f'<div class="podium">{"".join(kartu) or "<p>Belum ada kandidat.</p>"}</div>'
                      f'<p class="redup">Hasil per kandidat ditampilkan setelah '
                      f'voting ditutup.</p></section>')
        kelas_baris = "".join(
            f'<div class="kelas-baris"><span>{esc(r["kelas"] or "-")}</span>'
            f'<span class="kelas-angka">{r["sudah"]}/{r["total"]} sudah memilih</span></div>'
            for r in agg["per_kelas"]) or '<p class="redup">Belum ada data pemilih.</p>'
        diperbarui = f"{now:%H}.{now:%M} WIB"
        body = f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="30">
<title>Layar Hasil - {esc(nama_pemilu)}</title>
<link rel="stylesheet" href="/static/style.css"></head>
<body class="layar-body">
<div class="layar-grid">
  <div class="layar-kartu-atas">
    <div class="stat"><div class="stat-label">Total Pemilih</div>
      <div class="stat-angka">{agg['total_pemilih']}</div>
      <div class="stat-sub">siswa terdaftar</div></div>
    <div class="stat"><div class="stat-label">Suara Masuk</div>
      <div class="stat-angka">{agg['total_suara']}</div>
      <div class="stat-sub">suara yang masuk</div></div>
    <div class="stat"><div class="stat-label">Partisipasi</div>
      <div class="stat-angka">{fmt_persen(agg['partisipasi'])}</div>
      <div class="stat-sub">tingkat kehadiran</div></div>
    <div class="stat"><div class="stat-label">Hitung Mundur</div>
      {sisa_waktu_html(status, buka, tutup, now)}</div>
  </div>
  <div class="layar-tengah">{tengah}
    <section class="layar-panel"><h3>Partisipasi Pemilih</h3>
      <div class="donat-bungkus">{donat_svg(agg['partisipasi'])}
      <div class="donat-legend"><p><span class=" titik sudah"></span>Sudah Memilih: <b>{agg['sudah']}</b></p>
      <p><span class="titik belum"></span>Belum Memilih: <b>{agg['belum']}</b></p></div></div>
      <h3 class="sub">Partisipasi per Kelas</h3>{kelas_baris}</section>
  </div>
  <footer class="layar-kaki"><span>Status: {esc(status_teks)} &bull; Diperbarui: {diperbarui}</span>
  <span>{esc(nama_pemilu)}{(' &bull; ' + esc(nama_sekolah)) if nama_sekolah else ''} &bull; Realtime</span></footer>
</div>
</body></html>"""
        return self.send_html(body)


    # -----------------------------------------------------------------------
    # ADMIN: dashboard
    # -----------------------------------------------------------------------
    def admin_dashboard(self, user, flash, peringatan=""):
        agg = agregat()
        peng = agg["peng"]
        now = waktu_sekarang()
        status, buka, tutup = status_voting(peng, now)
        kelas_baris = "".join(
            f'<div class="bar-baris"><span class="bar-nama">{esc(r["kelas"] or "-")}</span>'
            f'<span class="bar-track"><span class="bar-isi toska" '
            f'style="width:{(r["sudah"] / r["total"] * 100) if r["total"] else 0:.1f}%"></span></span>'
            f'<span class="bar-angka">{r["sudah"]}/{r["total"]}</span></div>'
            for r in agg["per_kelas"]) or '<p class="redup">Belum ada data pemilih.</p>'
        body = f"""<meta http-equiv="refresh" content="30">
        <h2>Dashboard Utama</h2>
        <p class="redup">Ringkasan data pemilih dan hasil suara secara real-time
        &bull; {esc(peng['nama_pemilu'])}</p>
        <div class="stat-grid">
          <div class="stat"><div class="stat-label">Total Pemilih</div>
            <div class="stat-angka">{agg['total_pemilih']}</div>
            <div class="stat-sub">siswa terdaftar</div></div>
          <div class="stat"><div class="stat-label">Suara Masuk</div>
            <div class="stat-angka">{agg['total_suara']}</div>
            <div class="stat-sub">suara yang masuk</div></div>
          <div class="stat"><div class="stat-label">Tingkat Partisipasi</div>
            <div class="stat-angka">{fmt_persen(agg['partisipasi'])}</div>
            <div class="stat-sub">dari total pemilih</div></div>
          <div class="stat">{sisa_waktu_html(status, buka, tutup, now)}</div>
        </div>
        <section class="panel"><h3>Live Results — Podium 3 Kandidat Teratas</h3>
        {podium_html(agg['kandidat'], True)}
        <h3 class="sub">Grafik Suara per Kandidat</h3>
        {bar_kandidat_html(agg['kandidat'], True)}
        <p class="redup"><small>Total suara sah: {agg['total_suara']} &bull;
        Diperbarui otomatis tiap 30 detik &bull; {now:%H}.{now:%M} WIB</small></p></section>
        <div class="dua-kolom">
          <section class="panel"><h3>Partisipasi Pemilih</h3>
            <div class="donat-bungkus">{donat_svg(agg['partisipasi'])}
            <div class="donat-legend">
              <p><span class="titik sudah"></span>Sudah Memilih: <b>{agg['sudah']}</b></p>
              <p><span class="titik belum"></span>Belum Memilih: <b>{agg['belum']}</b></p>
            </div></div></section>
          <section class="panel"><h3>Partisipasi per Kelas</h3>{kelas_baris}</section>
        </div>"""
        return self.send_html(layout_admin(user, "Dashboard", body, flash,
                                           aktif="dashboard", peringatan=peringatan))

    # -----------------------------------------------------------------------
    # ADMIN: kandidat CRUD
    # -----------------------------------------------------------------------
    def admin_kandidat_list(self, user, qs, flash, peringatan=""):
        q = (qs.get("q") or [""])[0].strip()
        conn = db()
        if q:
            like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
            rows = conn.execute(
                """SELECT k.*, (SELECT COUNT(*) FROM suara s WHERE s.kandidat_id=k.id) suara
                   FROM kandidat k
                   WHERE k.nama LIKE ? ESCAPE '\\' OR k.kelas LIKE ? ESCAPE '\\'
                      OR CAST(k.nomor_urut AS TEXT) LIKE ? ESCAPE '\\'
                   ORDER BY k.nomor_urut""", (like, like, like)).fetchall()
        else:
            rows = conn.execute(
                """SELECT k.*, (SELECT COUNT(*) FROM suara s WHERE s.kandidat_id=k.id) suara
                   FROM kandidat k ORDER BY k.nomor_urut""").fetchall()
        conn.close()
        baris = ""
        for r in rows:
            foto = (f'<img class="thumb" src="/foto/{esc(r["foto"])}" alt="">'
                    if r["foto"] else '<span class="thumb kosong"></span>')
            baris += f"""<tr><td>{r['nomor_urut']:02d}</td><td>{foto}</td>
            <td><b>{esc(r['nama'])}</b></td><td>{esc(r['kelas'])}</td>
            <td>{esc(r['visi'])}</td><td>{r['suara']}</td>
            <td class="aksi"><a class="btn-kecil" href="/admin/kandidat/{r['id']}">Edit</a>
            <form method="post" action="/admin/kandidat/{r['id']}/hapus" class="inline"
              onsubmit="return confirm('Hapus kandidat {esc(r['nama'])}?')">
              <button type="submit" class="btn-kecil bahaya">Hapus</button></form></td></tr>"""
        kosong = '<tr><td colspan="7">Belum ada kandidat.</td></tr>' if not rows else ""
        body = f"""<h2>Manajemen Kandidat</h2>
        <p class="redup">Kelola data kandidat peserta {esc(_merek()['nama_pemilu'])}.</p>
        <p><a class="btn-emas" href="/admin/kandidat/baru">+ Tambah Kandidat</a></p>
        <form method="get" action="/admin/kandidat" class="cari">
          <input name="q" value="{esc(q)}" placeholder="Cari nama, kelas, atau no urut...">
          <button type="submit">Cari</button></form>
        <section class="panel"><table class="tabel">
          <thead><tr><th>No</th><th>Foto</th><th>Nama Kandidat</th><th>Kelas</th>
          <th>Visi Singkat</th><th>Suara</th><th>Aksi</th></tr></thead>
          <tbody>{baris}{kosong}</tbody></table></section>"""
        return self.send_html(layout_admin(user, "Kandidat", body, flash,
                                           aktif="kandidat", peringatan=peringatan))

    def admin_kandidat_form(self, user, kid, flash, peringatan=""):
        conn = db()
        k = None
        if kid is not None:
            k = conn.execute("SELECT * FROM kandidat WHERE id=?", (kid,)).fetchone()
            if not k:
                conn.close()
                return self.redirect("/admin/kandidat", "Data tidak ditemukan.")
        conn.close()
        judul = "Edit Kandidat" if k else "Tambah Kandidat"
        aksi = f"/admin/kandidat/{kid}/edit" if k else "/admin/kandidat/baru"
        foto_html = ""
        if k and k["foto"]:
            foto_html = f"""<div class="foto-saat-ini"><img src="/foto/{esc(k['foto'])}" alt="Foto kandidat">
            </div>
            <form method="post" action="/admin/kandidat/{kid}/foto-hapus" class="inline"
              onsubmit="return confirm('Hapus foto kandidat ini?')">
              <button type="submit" class="btn-kecil bahaya">Hapus Foto</button></form>"""
        body = f"""<h2>{judul}</h2>
        <section class="panel"><form method="post" action="{aksi}" enctype="multipart/form-data">
          <label>Nomor urut <input name="nomor_urut" type="number" min="1" step="1" required
            value="{k['nomor_urut'] if k else ''}"></label>
          <label>Nama kandidat <input name="nama" required value="{esc(k['nama']) if k else ''}"></label>
          <label>Kelas <input name="kelas" value="{esc(k['kelas']) if k else ''}"
            placeholder="mis. XI IPA 1"></label>
          <label>Visi singkat <input name="visi" value="{esc(k['visi']) if k else ''}"
            placeholder="mis. Sekolah bersih, ekskul aktif, suara siswa didengar"></label>
          {foto_html}
          <label>Foto kandidat (JPG/PNG/WebP, maks 5 MB)
            <input name="foto" type="file" accept=".jpg,.jpeg,.png,.webp"></label>
          <button type="submit" class="btn-emas">Simpan</button>
          <a class="btn-garis gelap" href="/admin/kandidat">Batal</a></form></section>"""
        return self.send_html(layout_admin(user, judul, body, flash,
                                           aktif="kandidat", peringatan=peringatan))

    def post_kandidat_simpan(self, user, kid, f, files):
        tujuan = f"/admin/kandidat/{kid}" if kid else "/admin/kandidat/baru"
        try:
            nomor = int((f.get("nomor_urut") or "").strip())
            if nomor <= 0:
                raise ValueError
        except ValueError:
            return self.redirect(tujuan, "Nomor urut harus bilangan bulat positif.")
        nama = (f.get("nama") or "").strip()
        if not nama:
            return self.redirect(tujuan, "Nama kandidat wajib diisi.")
        kelas = (f.get("kelas") or "").strip()
        visi = (f.get("visi") or "").strip()
        foto_baru, err = save_upload(files, "foto")
        if err:
            return self.redirect(tujuan, err)
        conn = db()
        bentrok = conn.execute(
            "SELECT id FROM kandidat WHERE nomor_urut=? AND id IS NOT ?",
            (nomor, kid)).fetchone()
        if bentrok:
            conn.close()
            if foto_baru:
                delete_upload(foto_baru)
            return self.redirect(tujuan, f"Nomor urut {nomor} sudah dipakai kandidat lain.")
        if kid is None:
            conn.execute(
                "INSERT INTO kandidat (nomor_urut, nama, kelas, visi, foto) VALUES (?,?,?,?,?)",
                (nomor, nama, kelas, visi, foto_baru))
            conn.commit()
            conn.close()
            return self.redirect("/admin/kandidat", f"Kandidat {nama} ditambahkan.")
        lama = conn.execute("SELECT foto FROM kandidat WHERE id=?", (kid,)).fetchone()
        if not lama:
            conn.close()
            if foto_baru:
                delete_upload(foto_baru)
            return self.redirect("/admin/kandidat", "Data tidak ditemukan.")
        foto = foto_baru or lama["foto"]
        conn.execute(
            "UPDATE kandidat SET nomor_urut=?, nama=?, kelas=?, visi=?, foto=? WHERE id=?",
            (nomor, nama, kelas, visi, foto, kid))
        conn.commit()
        conn.close()
        if foto_baru and lama["foto"] and lama["foto"] != foto_baru:
            delete_upload(lama["foto"])
        return self.redirect("/admin/kandidat", f"Kandidat {nama} disimpan.")

    def post_kandidat_hapus(self, user, kid):
        conn = db()
        k = conn.execute("SELECT * FROM kandidat WHERE id=?", (kid,)).fetchone()
        if not k:
            conn.close()
            return self.redirect("/admin/kandidat", "Data tidak ditemukan.")
        jml = conn.execute("SELECT COUNT(*) c FROM suara WHERE kandidat_id=?",
                           (kid,)).fetchone()["c"]
        if jml > 0:
            conn.close()
            return self.redirect(
                "/admin/kandidat",
                f"Kandidat {k['nama']} sudah memiliki {jml} suara dan TIDAK dapat "
                "dihapus. Kandidat bersuara hanya boleh diedit.")
        conn.execute("DELETE FROM kandidat WHERE id=?", (kid,))
        conn.commit()
        conn.close()
        delete_upload(k["foto"])
        return self.redirect("/admin/kandidat", f"Kandidat {k['nama']} dihapus.")

    def post_kandidat_foto_hapus(self, user, kid):
        conn = db()
        k = conn.execute("SELECT * FROM kandidat WHERE id=?", (kid,)).fetchone()
        if not k:
            conn.close()
            return self.redirect("/admin/kandidat", "Data tidak ditemukan.")
        conn.execute("UPDATE kandidat SET foto=NULL WHERE id=?", (kid,))
        conn.commit()
        conn.close()
        delete_upload(k["foto"])
        return self.redirect(f"/admin/kandidat/{kid}", "Foto kandidat dihapus.")

    # -----------------------------------------------------------------------
    # ADMIN: pemilih CRUD + impor massal
    # -----------------------------------------------------------------------
    def admin_pemilih_list(self, user, qs, flash, peringatan=""):
        q = (qs.get("q") or [""])[0].strip()
        conn = db()
        if q:
            like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
            rows = conn.execute(
                """SELECT * FROM pemilih
                   WHERE nis LIKE ? ESCAPE '\\' OR nama LIKE ? ESCAPE '\\'
                      OR kelas LIKE ? ESCAPE '\\'
                   ORDER BY kelas, nama LIMIT 300""", (like, like, like)).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM pemilih ORDER BY kelas, nama LIMIT 300").fetchall()
        conn.close()
        baris = ""
        for r in rows:
            lencana = ('<span class="lencana sudah">Sudah Memilih</span>' if r["sudah_memilih"]
                       else '<span class="lencana belum">Belum</span>')
            baris += f"""<tr><td><code>{esc(r['nis'])}</code></td>
            <td><b>{esc(r['nama'])}</b></td><td>{esc(r['kelas'])}</td><td>{lencana}</td>
            <td class="aksi"><a class="btn-kecil" href="/admin/pemilih/{r['id']}">Edit</a>
            <form method="post" action="/admin/pemilih/{r['id']}/hapus" class="inline"
              onsubmit="return confirm('Hapus pemilih {esc(r['nama'])}?')">
              <button type="submit" class="btn-kecil bahaya">Hapus</button></form></td></tr>"""
        kosong = '<tr><td colspan="5">Belum ada data pemilih.</td></tr>' if not rows else ""
        body = f"""<h2>Daftar Pemilih</h2>
        <p class="redup">Siswa yang berhak memilih. Admin hanya melihat status
        partisipasi — pilihan suara siswa TIDAK pernah terlihat.</p>
        <p><a class="btn-emas" href="/admin/pemilih/baru">+ Tambah Pemilih</a></p>
        <form method="get" action="/admin/pemilih" class="cari">
          <input name="q" value="{esc(q)}" placeholder="Cari NIS, nama, atau kelas...">
          <button type="submit">Cari</button></form>
        <section class="panel"><table class="tabel">
          <thead><tr><th>NIS</th><th>Nama</th><th>Kelas</th><th>Status</th><th>Aksi</th></tr></thead>
          <tbody>{baris}{kosong}</tbody></table></section>
        <section class="panel"><h3>Impor Pemilih Massal</h3>
        <p class="redup">Satu pemilih per baris, format: <code>NIS;Nama;Kelas</code>.
        Baris duplikat (NIS sudah ada) dan baris tidak valid dilewati.</p>
        <form method="post" action="/admin/pemilih/impor">
          <label>Data pemilih
            <textarea name="data" rows="7" placeholder="2026001;Andi Pratama;XI IPA 1&#10;2026002;Sari Lestari;XI IPA 1"></textarea></label>
          <button type="submit" class="btn-emas">Impor Sekarang</button></form></section>"""
        return self.send_html(layout_admin(user, "Pemilih", body, flash,
                                           aktif="pemilih", peringatan=peringatan))

    def admin_pemilih_form(self, user, pid, flash, peringatan=""):
        conn = db()
        p = None
        if pid is not None:
            p = conn.execute("SELECT * FROM pemilih WHERE id=?", (pid,)).fetchone()
            if not p:
                conn.close()
                return self.redirect("/admin/pemilih", "Data tidak ditemukan.")
        conn.close()
        judul = "Edit Pemilih" if p else "Tambah Pemilih"
        aksi = f"/admin/pemilih/{pid}/edit" if p else "/admin/pemilih/baru"
        nis_attr = " readonly" if (p and p["sudah_memilih"]) else ""
        catatan = ('<p class="redup">Pemilih ini SUDAH memilih: NIS tidak dapat '
                   'diubah dan data tidak dapat dihapus. Nama/kelas masih boleh '
                   'diperbaiki.</p>' if (p and p["sudah_memilih"]) else "")
        body = f"""<h2>{judul}</h2>
        <section class="panel">{catatan}
        <form method="post" action="{aksi}">
          <label>NIS <input name="nis" required value="{esc(p['nis']) if p else ''}"{nis_attr}></label>
          <label>Nama <input name="nama" required value="{esc(p['nama']) if p else ''}"></label>
          <label>Kelas <input name="kelas" value="{esc(p['kelas']) if p else ''}"
            placeholder="mis. XI IPA 1"></label>
          <button type="submit" class="btn-emas">Simpan</button>
          <a class="btn-garis gelap" href="/admin/pemilih">Batal</a></form></section>"""
        return self.send_html(layout_admin(user, judul, body, flash,
                                           aktif="pemilih", peringatan=peringatan))

    def post_pemilih_simpan(self, user, pid, f):
        tujuan = f"/admin/pemilih/{pid}" if pid else "/admin/pemilih/baru"
        nis = (f.get("nis") or "").strip()
        nama = (f.get("nama") or "").strip()
        kelas = (f.get("kelas") or "").strip()
        if not nis or not nama:
            return self.redirect(tujuan, "NIS dan nama wajib diisi.")
        conn = db()
        if pid is not None:
            lama = conn.execute("SELECT * FROM pemilih WHERE id=?", (pid,)).fetchone()
            if not lama:
                conn.close()
                return self.redirect("/admin/pemilih", "Data tidak ditemukan.")
            if lama["sudah_memilih"] and nis != lama["nis"]:
                conn.close()
                return self.redirect(
                    tujuan, "Pemilih yang sudah memilih tidak dapat diubah NIS-nya.")
        bentrok = conn.execute(
            "SELECT id FROM pemilih WHERE nis=? AND id IS NOT ?", (nis, pid)).fetchone()
        if bentrok:
            conn.close()
            return self.redirect(tujuan, f"NIS {nis} sudah terdaftar untuk pemilih lain.")
        if pid is None:
            conn.execute("INSERT INTO pemilih (nis, nama, kelas) VALUES (?,?,?)",
                         (nis, nama, kelas))
            pesan = f"Pemilih {nama} ditambahkan."
        else:
            conn.execute("UPDATE pemilih SET nis=?, nama=?, kelas=? WHERE id=?",
                         (nis, nama, kelas, pid))
            pesan = f"Pemilih {nama} disimpan."
        conn.commit()
        conn.close()
        return self.redirect("/admin/pemilih", pesan)

    def post_pemilih_hapus(self, user, pid):
        conn = db()
        p = conn.execute("SELECT * FROM pemilih WHERE id=?", (pid,)).fetchone()
        if not p:
            conn.close()
            return self.redirect("/admin/pemilih", "Data tidak ditemukan.")
        if p["sudah_memilih"]:
            conn.close()
            return self.redirect(
                "/admin/pemilih",
                f"{p['nama']} sudah menggunakan hak pilihnya dan TIDAK dapat dihapus.")
        conn.execute("DELETE FROM pemilih WHERE id=?", (pid,))
        conn.commit()
        conn.close()
        return self.redirect("/admin/pemilih", f"Pemilih {p['nama']} dihapus.")

    def post_pemilih_impor(self, user, f):
        mentah = f.get("data") or ""
        conn = db()
        ada = {r["nis"] for r in conn.execute("SELECT nis FROM pemilih")}
        tambah, lewati = 0, 0
        for baris in mentah.splitlines():
            baris = baris.strip()
            if not baris:
                continue
            bagian = [b.strip() for b in baris.split(";")]
            if len(bagian) != 3 or not bagian[0] or not bagian[1]:
                lewati += 1
                continue
            nis, nama, kelas = bagian
            if nis in ada:
                lewati += 1
                continue
            conn.execute("INSERT INTO pemilih (nis, nama, kelas) VALUES (?,?,?)",
                         (nis, nama, kelas))
            ada.add(nis)
            tambah += 1
        conn.commit()
        conn.close()
        return self.redirect(
            "/admin/pemilih",
            f"Impor selesai: {tambah} pemilih ditambahkan, {lewati} baris dilewati.")

    # -----------------------------------------------------------------------
    # ADMIN: hasil suara + ekspor CSV
    # -----------------------------------------------------------------------
    def admin_hasil(self, user, flash, peringatan=""):
        agg = agregat()
        total = max(1, agg["total_suara"])
        baris = ""
        for k in agg["kandidat"]:
            baris += (f"<tr><td>{k['nomor_urut']:02d}</td><td><b>{esc(k['nama'])}</b></td>"
                      f"<td>{esc(k['kelas'])}</td><td>{k['suara']}</td>"
                      f"<td>{fmt_persen(k['suara'] / total * 100)}</td></tr>")
        kelas_baris = ""
        for r in agg["per_kelas"]:
            p = (r["sudah"] / r["total"] * 100) if r["total"] else 0.0
            kelas_baris += (f"<tr><td>{esc(r['kelas'] or '-')}</td><td>{r['total']}</td>"
                            f"<td>{r['sudah']}</td><td>{r['total'] - r['sudah']}</td>"
                            f"<td>{fmt_persen(p)}</td></tr>")
        body = f"""<h2>Hasil Suara</h2>
        <p class="redup">{esc(agg['peng']['nama_pemilu'])} &bull; total suara sah:
        <b>{agg['total_suara']}</b> dari {agg['total_pemilih']} pemilih
        ({fmt_persen(agg['partisipasi'])}).</p>
        <p><a class="btn-emas" href="/admin/hasil/hasil.xlsx">Unduh Excel (.xlsx)</a>
        <a class="btn-garis gelap" href="/admin/hasil/hasil.csv">Unduh CSV Hasil Akhir</a>
        <a class="btn-garis gelap" href="/admin/hasil/partisipasi.csv">Unduh CSV Partisipasi</a></p>
        <p class="redup"><small>Berkas Excel hasil akhir dapat dipakai sebagai
        bukti resmi: di dalamnya tercantum pemenang, ringkasan partisipasi,
        waktu penutupan voting, dan waktu laporan dibuat.</small></p>
        <p><a class="btn-garis gelap" href="/admin/berita-acara">Berita Acara Hasil (Dokumen Resmi)</a></p>
        <p class="redup"><small>Berita Acara adalah dokumen resmi hasil pemilihan
        untuk dicetak / disimpan sebagai PDF dan ditandatangani panitia.</small></p>
        <section class="panel"><h3>Perolehan Suara per Kandidat</h3>
        <table class="tabel"><thead><tr><th>No</th><th>Nama</th><th>Kelas</th>
        <th>Suara</th><th>Persen</th></tr></thead>
        <tbody>{baris or '<tr><td colspan="5">Belum ada kandidat.</td></tr>'}</tbody></table></section>
        <section class="panel"><h3>Partisipasi per Kelas</h3>
        <table class="tabel"><thead><tr><th>Kelas</th><th>Total Pemilih</th>
        <th>Sudah</th><th>Belum</th><th>Persen</th></tr></thead>
        <tbody>{kelas_baris or '<tr><td colspan="5">Belum ada data pemilih.</td></tr>'}</tbody></table></section>"""
        return self.send_html(layout_admin(user, "Hasil Suara", body, flash,
                                           aktif="hasil", peringatan=peringatan))

    def _nama_csv(self, awalan):
        peng = _merek()
        tgl = waktu_sekarang().strftime("%Y-%m-%d")
        return f"{awalan}-{slug_nama(peng['nama_pemilu'])}-{tgl}.csv"

    def csv_hasil(self, user):
        agg = agregat()
        total = max(1, agg["total_suara"])
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["Nomor Urut", "Nama Kandidat", "Kelas", "Suara", "Persen"])
        for k in agg["kandidat"]:
            w.writerow([k["nomor_urut"], k["nama"], k["kelas"], k["suara"],
                        fmt_persen(k["suara"] / total * 100)])
        w.writerow(["", "TOTAL", "", agg["total_suara"], ""])
        data = ("\ufeff" + buf.getvalue()).encode("utf-8")
        return self.send_bytes(data, "text/csv; charset=utf-8",
                               self._nama_csv("hasil"))

    def csv_partisipasi(self, user):
        agg = agregat()
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["Kelas", "Total Pemilih", "Sudah Memilih", "Belum Memilih", "Persen"])
        for r in agg["per_kelas"]:
            p = (r["sudah"] / r["total"] * 100) if r["total"] else 0.0
            w.writerow([r["kelas"], r["total"], r["sudah"],
                        r["total"] - r["sudah"], fmt_persen(p)])
        w.writerow(["TOTAL", agg["total_pemilih"], agg["sudah"], agg["belum"],
                    fmt_persen(agg["partisipasi"])])
        data = ("\ufeff" + buf.getvalue()).encode("utf-8")
        return self.send_bytes(data, "text/csv; charset=utf-8",
                               self._nama_csv("partisipasi"))

    # -----------------------------------------------------------------------
    # ADMIN: ekspor Excel (.xlsx) — laporan hasil akhir sebagai BUKTI resmi.
    # Gaya mengikuti ekspor Excel aplikasi Absen: blok judul, baris header
    # gelap, garis tepi tipis, freeze panes, landscape fit-to-width.
    # Kedua sheet DILINDUNGI (worksheet protection) supaya berkas bukti
    # tidak mudah diubah — ini pencegah suntingan kasual, BUKAN enkripsi.
    # Satu-satunya dependensi eksternal aplikasi ini: openpyxl.
    # -----------------------------------------------------------------------
    def _nama_file_excel(self):
        peng = _merek()
        bersih = re.sub(r"[^\w\-]+", "-", peng["nama_pemilu"]).strip("-") or "Pemilu"
        return f"Hasil-{bersih}.xlsx"

    def excel_hasil(self, user):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        agg = agregat()
        peng = agg["peng"]
        now = waktu_sekarang()
        status, buka, tutup = status_voting(peng, now)
        total = max(1, agg["total_suara"])
        nama_pemilu = peng["nama_pemilu"] or "Pemilihan Ketua OSIS"
        nama_sekolah = peng["nama_sekolah"]

        tipis = Side(style="thin", color="9AA3B2")
        border = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
        fill_head = PatternFill("solid", fgColor="203764")
        font_head = Font(bold=True, color="FFFFFF", size=11)
        fill_menang = PatternFill("solid", fgColor="F2C14E")
        font_menang = Font(bold=True, color="1F3864", size=11)

        wb = Workbook()

        # ---- Sheet 1: Hasil Akhir -------------------------------------
        ws = wb.active
        ws.title = "Hasil Akhir"
        nkol = 6
        akhir = get_column_letter(nkol)
        off = 1 if nama_sekolah else 0  # baris nama sekolah di paling atas
        if nama_sekolah:
            ws.merge_cells(f"A1:{akhir}1")
            c = ws["A1"]; c.value = nama_sekolah
            c.font = Font(bold=True, size=16, color="1F3864")
        ws.merge_cells(f"A{1 + off}:{akhir}{1 + off}")
        c = ws[f"A{1 + off}"]; c.value = nama_pemilu
        c.font = Font(bold=True, size=15, color="1F3864")
        ws.merge_cells(f"A{2 + off}:{akhir}{2 + off}")
        c = ws[f"A{2 + off}"]; c.value = "Hasil Akhir Perolehan Suara"
        c.font = Font(bold=True, size=11, color="404040")
        header = ["No Urut", "Nama", "Kelas", "Jumlah Suara", "Persentase", "Keterangan"]
        hr = 4 + off
        for j, h in enumerate(header, start=1):
            c = ws.cell(row=hr, column=j, value=h)
            c.fill = fill_head; c.font = font_head
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        r = hr + 1
        for i, k in enumerate(agg["kandidat"]):
            pemenang = (i == 0 and k["suara"] > 0)
            vals = [k["nomor_urut"], k["nama"], k["kelas"], k["suara"],
                    fmt_persen(k["suara"] / total * 100),
                    "Pemenang" if pemenang else ""]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = border
                c.alignment = Alignment(
                    horizontal="left" if j == 2 else "center", vertical="center")
                if pemenang:
                    c.fill = fill_menang; c.font = font_menang
            r += 1
        if not agg["kandidat"]:
            ws.cell(row=r, column=1, value="Belum ada kandidat.")
            r += 1
        # Ringkasan di bawah tabel.
        r += 1
        ringkasan = [
            ("Total Pemilih Terdaftar", agg["total_pemilih"]),
            ("Suara Masuk", agg["total_suara"]),
            ("Belum Memilih", agg["belum"]),
            ("Tingkat Partisipasi", fmt_persen(agg["partisipasi"])),
        ]
        for label, nilai in ringkasan:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nkol)
            c = ws.cell(row=r, column=1, value=f"{label}: {nilai}")
            c.font = Font(bold=True, size=11, color="1F3864")
            r += 1
        # Baris BUKTI: kapan voting ditutup & kapan laporan ini dibuat.
        r += 1
        if status == "selesai":
            baris_tutup = f"Voting ditutup: {fmt_jadwal(tutup)}"
        else:
            baris_tutup = "Voting masih berlangsung"
        for teks in (baris_tutup, f"Laporan dibuat: {fmt_jadwal(now)}"):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nkol)
            c = ws.cell(row=r, column=1, value=teks)
            c.font = Font(italic=True, size=10, color="404040")
            r += 1
        for j, wdt in {1: 10, 2: 32, 3: 16, 4: 14, 5: 12, 6: 14}.items():
            ws.column_dimensions[get_column_letter(j)].width = wdt
        ws.freeze_panes = f"A{hr + 1}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        # Kunci sheet: berkas bukti tidak bisa disunting tanpa membuka kunci.
        ws.protection.sheet = True
        ws.protection.password = KUNCI_EXCEL

        # ---- Sheet 2: Partisipasi Kelas --------------------------------
        ws2 = wb.create_sheet("Partisipasi Kelas")
        ws2.merge_cells("A1:E1")
        c = ws2["A1"]; c.value = f"Partisipasi per Kelas — {nama_pemilu}"
        c.font = Font(bold=True, size=13, color="1F3864")
        header2 = ["Kelas", "Total Pemilih", "Sudah Memilih", "Belum Memilih",
                   "Partisipasi (%)"]
        for j, h in enumerate(header2, start=1):
            c = ws2.cell(row=3, column=j, value=h)
            c.fill = fill_head; c.font = font_head
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        r = 4
        for baris_kelas in agg["per_kelas"]:
            p = ((baris_kelas["sudah"] / baris_kelas["total"] * 100)
                 if baris_kelas["total"] else 0.0)
            vals = [baris_kelas["kelas"], baris_kelas["total"], baris_kelas["sudah"],
                    baris_kelas["total"] - baris_kelas["sudah"], fmt_persen(p)]
            for j, v in enumerate(vals, start=1):
                c = ws2.cell(row=r, column=j, value=v)
                c.border = border
                c.alignment = Alignment(
                    horizontal="left" if j == 1 else "center", vertical="center")
            r += 1
        vals = ["TOTAL", agg["total_pemilih"], agg["sudah"], agg["belum"],
                fmt_persen(agg["partisipasi"])]
        for j, v in enumerate(vals, start=1):
            c = ws2.cell(row=r, column=j, value=v)
            c.font = Font(bold=True)
            c.border = border
            c.alignment = Alignment(
                horizontal="left" if j == 1 else "center", vertical="center")
        for j, wdt in {1: 22, 2: 15, 3: 15, 4: 15, 5: 16}.items():
            ws2.column_dimensions[get_column_letter(j)].width = wdt
        ws2.freeze_panes = "A4"
        ws2.page_setup.orientation = "landscape"
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 0
        ws2.sheet_properties.pageSetUpPr.fitToPage = True
        ws2.protection.sheet = True
        ws2.protection.password = KUNCI_EXCEL

        buf = io.BytesIO()
        wb.save(buf)
        return self.send_bytes(
            buf.getvalue(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            self._nama_file_excel())

    # -----------------------------------------------------------------------
    # ADMIN: Berita Acara Hasil — dokumen resmi siap cetak (halaman putih
    # polos TANPA kerangka admin; tombol cetak hilang saat dicetak).
    # Admin mencetak / menyimpan sebagai PDF dari browser — salinan PDF
    # itulah bukti resmi yang tidak bisa disunting.
    # -----------------------------------------------------------------------
    def berita_acara(self, user):
        agg = agregat()
        peng = agg["peng"]
        now = waktu_sekarang()
        status, buka, tutup = status_voting(peng, now)
        total = max(1, agg["total_suara"])
        nama_pemilu = peng["nama_pemilu"] or "Pemilihan Ketua OSIS"
        nama_sekolah = peng["nama_sekolah"]
        baris = ""
        for i, k in enumerate(agg["kandidat"]):
            ket = " <b>(Pemenang)</b>" if (i == 0 and k["suara"] > 0) else ""
            baris += (f"<tr><td class='tengah'>{k['nomor_urut']}</td>"
                      f"<td>{esc(k['nama'])}{ket}</td><td>{esc(k['kelas'])}</td>"
                      f"<td class='tengah'>{k['suara']}</td>"
                      f"<td class='tengah'>{fmt_persen(k['suara'] / total * 100)}</td></tr>")
        if not agg["kandidat"]:
            baris = '<tr><td colspan="5">Belum ada kandidat.</td></tr>'
        if status == "selesai":
            baris_tutup = fmt_jadwal(tutup)
        elif status == "belum-diatur":
            baris_tutup = "Jadwal belum diatur"
        elif status == "belum-dibuka":
            baris_tutup = f"Belum ditutup (dibuka {fmt_jadwal(buka)})"
        else:
            baris_tutup = f"Belum ditutup (rencana {fmt_jadwal(tutup)})"
        kepala_sekolah = f"<p>{esc(nama_sekolah)}</p>" if nama_sekolah else ""
        body = f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Berita Acara Hasil - {esc(nama_pemilu)}</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: Georgia, 'Times New Roman', serif; margin: 0;
         background: #e8e8e8; color: #111; }}
  .lembar {{ background: #fff; max-width: 820px; margin: 24px auto;
            padding: 48px 56px; box-shadow: 0 2px 14px rgba(0,0,0,.25); }}
  .kop {{ text-align: center; border-bottom: 3px double #111;
         padding-bottom: 14px; margin-bottom: 22px; }}
  .kop .sekolah {{ font-size: 1.35rem; font-weight: bold; letter-spacing: .04em; }}
  .kop .pemilu {{ font-size: 1.05rem; margin-top: 2px; }}
  h1 {{ text-align: center; font-size: 1.3rem; letter-spacing: .06em;
       margin: 6px 0 4px; text-decoration: underline; }}
  .sub {{ text-align: center; margin: 0 0 20px; }}
  table {{ width: 100%; border-collapse: collapse; margin: 14px 0 20px; }}
  th, td {{ border: 1px solid #333; padding: 7px 10px; font-size: .98rem; }}
  th {{ background: #f0f0f0; }}
  td.tengah {{ text-align: center; }}
  .ringkasan {{ margin: 18px 0; }}
  .ringkasan td {{ border: 0; padding: 3px 8px 3px 0; }}
  .ttd {{ display: flex; gap: 40px; margin-top: 44px; }}
  .ttd > div {{ flex: 1; text-align: center; }}
  .ttd .ruang {{ height: 84px; }}
  .ttd .nama {{ border-top: 1px solid #111; display: inline-block;
               min-width: 200px; padding-top: 4px; }}
  .btn-cetak {{ display: block; margin: 18px auto; padding: 12px 26px;
    font-size: 1rem; font-family: inherit; cursor: pointer;
    background: #1d6fd1; color: #fff; border: 0; border-radius: 8px; }}
  @media print {{
    body {{ background: #fff; }}
    .lembar {{ box-shadow: none; margin: 0; max-width: none; padding: 12px; }}
    .btn-cetak, .tanpa-cetak {{ display: none !important; }}
  }}
</style></head>
<body>
<button class="btn-cetak" onclick="window.print()">Cetak / Simpan PDF</button>
<div class="lembar">
  <div class="kop">
    <div class="sekolah">{esc(nama_sekolah) if nama_sekolah else 'SEKOLAH'}</div>
    <div class="pemilu">{esc(nama_pemilu)}</div>
  </div>
  <h1>BERITA ACARA HASIL PEMILIHAN</h1>
  <p class="sub">Panitia {esc(nama_pemilu)} menyatakan hasil pemungutan suara
  elektronik (e-Voting) sebagai berikut:</p>
  <table>
    <thead><tr><th>No. Urut</th><th>Nama Kandidat</th><th>Kelas</th>
    <th>Jumlah Suara</th><th>Persentase</th></tr></thead>
    <tbody>{baris}</tbody>
  </table>
  <table class="ringkasan">
    <tr><td>Total Pemilih Terdaftar</td><td>: <b>{agg['total_pemilih']}</b></td></tr>
    <tr><td>Suara Masuk</td><td>: <b>{agg['total_suara']}</b></td></tr>
    <tr><td>Belum Memilih</td><td>: <b>{agg['belum']}</b></td></tr>
    <tr><td>Tingkat Partisipasi</td><td>: <b>{fmt_persen(agg['partisipasi'])}</b></td></tr>
    <tr><td>Waktu Voting Ditutup</td><td>: <b>{esc(baris_tutup)}</b></td></tr>
    <tr><td>Laporan Dibuat</td><td>: <b>{fmt_jadwal(now)}</b></td></tr>
  </table>
  <p>Demikian berita acara ini dibuat dengan sebenarnya untuk digunakan
  sebagaimana mestinya. Suara pemilih bersifat rahasia; dokumen ini hanya
  memuat hasil agregat.</p>
  <div class="ttd">
    <div>
      <p>Ketua Panitia</p>
      <div class="ruang"></div>
      <span class="nama">( Nama Lengkap )</span>
    </div>
    <div>
      <p>Mengetahui,<br>Kepala Sekolah</p>
      <div class="ruang"></div>
      <span class="nama">( Nama Lengkap )</span>
    </div>
  </div>
  {kepala_sekolah}
</div>
<p class="tanpa-cetak" style="text-align:center">
  <a href="/admin/hasil">Kembali ke Hasil Suara</a></p>
</body></html>"""
        return self.send_html(body)

    # -----------------------------------------------------------------------
    # ADMIN: pengaturan
    # -----------------------------------------------------------------------
    def admin_pengaturan(self, user, flash, peringatan=""):
        peng = _merek()
        logo_html_ini = ""
        if peng["logo"]:
            logo_html_ini = f"""<div class="foto-saat-ini"><img src="/foto/{esc(peng['logo'])}" alt="Logo saat ini"></div>
            <form method="post" action="/admin/pengaturan/logo-hapus" class="inline"
              onsubmit="return confirm('Hapus logo? Brand kembali ke ikon kotak suara bawaan.')">
              <button type="submit" class="btn-kecil bahaya">Hapus Logo</button></form>"""
        cek_live = " checked" if peng["hasil_live"] == "1" else ""
        body = f"""<h2>Pengaturan</h2>
        <section class="panel"><h3>Identitas &amp; Jadwal</h3>
        <form method="post" action="/admin/pengaturan">
          <label>Nama pemilihan <input name="nama_pemilu" required
            value="{esc(peng['nama_pemilu'])}"></label>
          <label>Nama sekolah <input name="nama_sekolah"
            value="{esc(peng['nama_sekolah'])}" placeholder="mis. SMA Negeri 1 Contoh"></label>
          <label>Waktu buka voting <input name="waktu_buka" type="datetime-local"
            value="{esc(peng['waktu_buka'])}"></label>
          <label>Waktu tutup voting <input name="waktu_tutup" type="datetime-local"
            value="{esc(peng['waktu_tutup'])}"></label>
          <label class="cek"><input type="checkbox" name="hasil_live" value="1"{cek_live}>
            Tampilkan hasil per kandidat secara langsung (live) di Layar Tampil
            selama voting berlangsung. Bila dimatikan, hasil baru tampil setelah
            voting ditutup; partisipasi tetap terlihat.</label>
          <button type="submit" class="btn-emas">Simpan Pengaturan</button></form></section>
        <section class="panel"><h3>Logo Sekolah</h3>
        <p class="redup">Logo tampil di panel admin dan kepala halaman pemilih.
        Kosongkan untuk ikon kotak suara bawaan.</p>
        {logo_html_ini}
        <form method="post" action="/admin/pengaturan/logo" enctype="multipart/form-data">
          <label>Unggah / ganti logo <input name="logo" type="file"
            accept=".jpg,.jpeg,.png,.webp" required></label>
          <button type="submit" class="btn-emas">Unggah Logo</button></form></section>"""
        return self.send_html(layout_admin(user, "Pengaturan", body, flash,
                                           aktif="pengaturan", peringatan=peringatan))

    def post_pengaturan(self, user, f):
        nama_pemilu = (f.get("nama_pemilu") or "").strip()
        if not nama_pemilu:
            return self.redirect("/admin/pengaturan", "Nama pemilihan wajib diisi.")
        buka = (f.get("waktu_buka") or "").strip()
        tutup = (f.get("waktu_tutup") or "").strip()
        dbuka, dtutup = parse_jadwal(buka), parse_jadwal(tutup)
        if (buka and dbuka is None) or (tutup and dtutup is None):
            return self.redirect("/admin/pengaturan", "Format jadwal tidak valid.")
        if dbuka and dtutup and dtutup <= dbuka:
            return self.redirect("/admin/pengaturan",
                                 "Waktu tutup harus sesudah waktu buka.")
        conn = db()
        set_setting(conn, "nama_pemilu", nama_pemilu)
        set_setting(conn, "nama_sekolah", (f.get("nama_sekolah") or "").strip())
        set_setting(conn, "waktu_buka", buka)
        set_setting(conn, "waktu_tutup", tutup)
        set_setting(conn, "hasil_live", "1" if f.get("hasil_live") == "1" else "0")
        conn.commit()
        conn.close()
        return self.redirect("/admin/pengaturan", "Pengaturan disimpan.")

    def post_logo(self, user, files):
        nama, err = save_upload(files, "logo")
        if err:
            return self.redirect("/admin/pengaturan", err)
        if not nama:
            return self.redirect("/admin/pengaturan", "Pilih berkas logo dulu.")
        conn = db()
        lama = get_setting(conn, "logo", "")
        set_setting(conn, "logo", nama)
        conn.commit()
        conn.close()
        if lama and lama != nama:
            delete_upload(lama)
        return self.redirect("/admin/pengaturan", "Logo disimpan.")

    def post_logo_hapus(self, user):
        conn = db()
        lama = get_setting(conn, "logo", "")
        set_setting(conn, "logo", "")
        conn.commit()
        conn.close()
        delete_upload(lama)
        return self.redirect("/admin/pengaturan", "Logo dihapus.")


    # -----------------------------------------------------------------------
    # ADMIN: backup & pulihkan
    # -----------------------------------------------------------------------
    def admin_backup(self, user, flash, peringatan=""):
        body = """<h2>Backup &amp; Pulihkan</h2>
        <section class="panel"><h3>Unduh Backup</h3>
        <p><small>Satu berkas ZIP berisi SELURUH data aplikasi ini: database
        (semua kandidat, pemilih, suara, akun admin, dan pengaturan) serta
        semua berkas di folder unggahan (foto kandidat dan logo). Simpan
        berkas ini di tempat yang aman — berkas inilah yang dipakai untuk
        memulihkan data bila suatu saat data di server hilang.</small></p>
        <p><a href="/admin/backup/unduh"><button type="button">Unduh Backup (.zip)</button></a></p></section>
        <section class="panel"><h3>Pulihkan dari Backup</h3>
        <div class="flash warn">PERINGATAN: memulihkan backup akan MENGGANTI
        SELURUH data saat ini (semua kandidat, pemilih, suara, foto, dan
        pengaturan) dengan isi berkas backup. Data yang ada sekarang akan
        hilang. Bila ragu, unduh backup terbaru dulu sebelum memulihkan.</div>
        <form method="post" action="/admin/backup/pulihkan" enctype="multipart/form-data"
          onsubmit="return confirm('Pulihkan backup ini? SELURUH data saat ini akan diganti dengan isi backup.')">
          <label>Berkas backup (.zip) <input name="berkas" type="file" accept=".zip" required></label>
          <button type="submit" class="danger">Pulihkan Backup</button></form></section>"""
        return self.send_html(layout_admin(user, "Backup & Pulihkan", body, flash,
                                           aktif="backup", peringatan=peringatan))

    def admin_backup_unduh(self, user):
        """Bangun ZIP backup saat itu juga: salinan database yang konsisten
        (lewat sqlite3 backup API), semua berkas di folder uploads di bawah
        uploads/, dan manifest.json penanda aplikasi."""
        tmpdir = tempfile.mkdtemp(prefix="vote-backup-")
        try:
            db_copy = os.path.join(tmpdir, "vote.db")
            src = sqlite3.connect(DB_PATH)
            try:
                dst = sqlite3.connect(db_copy)
                try:
                    src.backup(dst)
                finally:
                    dst.close()
            finally:
                src.close()
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(db_copy, "vote.db")
                if os.path.isdir(UPLOADS_DIR):
                    for name in sorted(os.listdir(UPLOADS_DIR)):
                        fpath = os.path.join(UPLOADS_DIR, name)
                        if os.path.isfile(fpath):
                            z.write(fpath, "uploads/" + name)
                manifest = {
                    "app": "vote-osis",
                    "dibuat": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                    "versi": 1,
                }
                z.writestr("manifest.json", json.dumps(manifest, indent=2))
            data = buf.getvalue()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        fname = "backup-vote-" + datetime.date.today().isoformat() + ".zip"
        return self.send_bytes(data, "application/zip", fname)

    def post_backup_pulihkan(self, user, files):
        """Ganti database + isi folder uploads dengan isi berkas backup.
        Semua validasi dilakukan SEBELUM ada berkas yang disentuh; bila
        berkas ditolak, data saat ini tidak berubah sama sekali."""
        back = "/admin/backup"
        bukan = "Berkas ini bukan backup aplikasi PEMILU OSIS."
        if "berkas" not in files or not files["berkas"][0]:
            return self.redirect(back, "Pilih berkas backup (.zip) dulu.")
        _fname, data = files["berkas"]
        try:
            zf = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            return self.redirect(back, "Berkas tidak valid — bukan file ZIP backup.")
        names = set(zf.namelist())
        if "manifest.json" not in names:
            return self.redirect(back, bukan)
        try:
            manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        except Exception:
            return self.redirect(back, bukan)
        if not isinstance(manifest, dict) or manifest.get("app") != "vote-osis":
            return self.redirect(back, bukan)
        if "vote.db" not in names:
            return self.redirect(back, bukan)
        tmpdir = tempfile.mkdtemp(prefix="vote-pulih-")
        try:
            db_baru = os.path.join(tmpdir, "vote.db")
            try:
                with open(db_baru, "wb") as fh:
                    fh.write(zf.read("vote.db"))
                uploads_baru = os.path.join(tmpdir, "uploads")
                os.makedirs(uploads_baru, exist_ok=True)
                for name in zf.namelist():
                    if not name.startswith("uploads/"):
                        continue
                    base = os.path.basename(name)
                    if not base or not FOTO_NAME_RE.fullmatch(base):
                        continue  # lewati entri asing / upaya zip-slip
                    with open(os.path.join(uploads_baru, base), "wb") as fh:
                        fh.write(zf.read(name))
            except (zipfile.BadZipFile, RuntimeError, OSError):
                return self.redirect(back, "Berkas tidak valid — isi backup rusak.")
            # Pastikan database di dalam backup utuh & memang milik aplikasi ini.
            try:
                cek = sqlite3.connect(db_baru)
                tabel = {r[0] for r in cek.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'")}
                cek.close()
            except sqlite3.Error:
                return self.redirect(back, "Berkas tidak valid — database di dalam backup rusak.")
            if not {"users", "kandidat", "pemilih", "suara", "settings"} <= tabel:
                return self.redirect(back, "Berkas tidak valid — database di dalam backup tidak lengkap.")
            # --- Semua validasi lolos: titik tukar data. ---
            db_dir = os.path.dirname(DB_PATH) or "."
            os.makedirs(db_dir, exist_ok=True)
            if os.path.isfile(DB_PATH):
                # Salinan pengaman data SEBELUM dipulihkan, di samping DB aktif.
                shutil.copyfile(DB_PATH, os.path.join(db_dir, "vote-sebelum-pulih.db"))
            os.makedirs(UPLOADS_DIR, exist_ok=True)
            for lama in os.listdir(UPLOADS_DIR):
                p = os.path.join(UPLOADS_DIR, lama)
                if os.path.isfile(p):
                    os.remove(p)
            for name in os.listdir(uploads_baru):
                shutil.copyfile(os.path.join(uploads_baru, name),
                                os.path.join(UPLOADS_DIR, name))
            shutil.copyfile(db_baru, DB_PATH)
            # Backup dari versi lama mungkin belum punya tabel baru:
            # jalankan skema pada database hasil pulihkan.
            conn = db()
            conn.executescript(SCHEMA)
            conn.commit()
            conn.close()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        # Database berganti (termasuk akun admin di dalamnya): semua sesi
        # login lama tidak berlaku lagi, admin masuk ulang dari /login.
        LOGIN_SESSIONS.clear()
        return self.redirect("/login", "Backup dipulihkan. Silakan masuk lagi.")


def main():
    init_db()
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"PEMILU OSIS jalan di http://127.0.0.1:{PORT}  (DB: {DB_PATH}, uploads: {UPLOADS_DIR})")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nBerhenti.")


if __name__ == "__main__":
    main()
