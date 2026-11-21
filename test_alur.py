#!/usr/bin/env python3
"""Uji alur end-to-end aplikasi PEMILU OSIS — e-Voting Sekolah lewat HTTP.

Skrip ini MENJALANKAN SENDIRI server aplikasi pada port uji dengan
database & folder unggahan SEMENTARA (di /tmp), lalu memeriksanya:

  python3 test_alur.py

Fase-fasenya memakai VOTE_NOW_OVERRIDE untuk mengendalikan "sekarang":
  A) sebelum voting dibuka — CRUD admin, impor, penolakan-penolakan awal
  B) voting berlangsung — alur memilih penuh, suara ganda ditolak,
     kerahasiaan suara, integritas data, hasil live, CSV, backup
  C) sesudah voting ditutup — Excel bukti (terkunci + baris bukti),
     Berita Acara, ganti password
  D) server BARU kosong — pulihkan backup dari fase B, data kembali utuh
"""
import io, json, os, re, sqlite3, subprocess, sys, tempfile, time
import urllib.request, urllib.parse, urllib.error, http.cookiejar, uuid, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("TEST_PORT", "8793"))
BASE = f"http://127.0.0.1:{PORT}"

# PNG 1x1 piksel (berkas gambar uji terkecil)
PNG_1PX = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000a49444154789c6360000000020001e221bc330000000049454e44ae426082")

JUMLAH = [0]


def client():
    cj = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))


def req(op, path, data=None):
    url = BASE + path
    body = urllib.parse.urlencode(data).encode() if data is not None else None
    r = op.open(urllib.request.Request(url, data=body))
    return r.status, r.geturl(), r.read().decode("utf-8", "replace")


def req_multipart(op, path, fields, files):
    """files: nama_field -> (nama_berkas, bytes, content_type)"""
    boundary = "----uji" + uuid.uuid4().hex
    chunks = []
    for name, value in fields.items():
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
        chunks.append(str(value).encode() + b"\r\n")
    for name, (fname, data, ctype) in files.items():
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append((f'Content-Disposition: form-data; name="{name}"; '
                       f'filename="{fname}"\r\n').encode())
        chunks.append(f"Content-Type: {ctype}\r\n\r\n".encode())
        chunks.append(data + b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode())
    body = b"".join(chunks)
    r = op.open(urllib.request.Request(
        BASE + path, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}))
    return r.status, r.geturl(), r.read().decode("utf-8", "replace")


def req_bytes(op, path):
    r = op.open(BASE + path)
    return r.status, r.geturl(), r.headers, r.read()


def must(cond, msg):
    JUMLAH[0] += 1
    print(("OK  " if cond else "GAGAL ") + msg)
    if not cond:
        sys.exit(1)


def mulai(tmp, now_override=None):
    env = dict(os.environ)
    env["VOTE_DB"] = os.path.join(tmp, "sub", "vote.db")
    env["VOTE_UPLOADS"] = os.path.join(tmp, "uploads")
    env["PORT"] = str(PORT)
    if now_override:
        env["VOTE_NOW_OVERRIDE"] = now_override
    proc = subprocess.Popen(
        [sys.executable, os.path.join(HERE, "app.py")],
        env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    siap = False
    for _ in range(60):
        try:
            urllib.request.urlopen(BASE + "/login", timeout=2)
            siap = True
            break
        except Exception:
            time.sleep(0.2)
    must(siap, "server uji menyala")
    return proc


def berhenti(proc):
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except Exception:
        proc.kill()
    time.sleep(0.3)


def login_admin(op, password="admin123"):
    s, url, _ = req(op, "/login", {"username": "admin", "password": password})
    return url


PENGATURAN = {"nama_pemilu": "Pemilihan Ketua OSIS 2026",
              "nama_sekolah": "SMA Negeri 1 Contoh",
              "waktu_buka": "2026-11-10T08:00",
              "waktu_tutup": "2026-11-10T16:00",
              "hasil_live": "1"}


def main():
    tmp = tempfile.mkdtemp(prefix="vote-test-")
    db_path = os.path.join(tmp, "sub", "vote.db")

    # =====================================================================
    # FASE A — sebelum voting dibuka (sekarang 07:00, buka 08:00)
    # =====================================================================
    proc = mulai(tmp, "2026-11-10T07:00:00")
    try:
        publik = client()
        s, url, home = req(publik, "/")
        must(s == 200 and "Cek Identitas Pemilih" in home and "Verifikasi NIS" in home,
             "GET / -> halaman pemilih terang dengan kartu Cek Identitas Pemilih")
        s, url, js = req(publik, "/static/app.js")
        must("pw-toggle" in js and 'input[type="password"]' in js,
             "app.js memuat logika toggle mata password (pw-toggle)")
        s, url, css = req(publik, "/static/style.css")
        must(".pw-toggle" in css and ".pw-wrap" in css,
             "style.css memuat gaya .pw-toggle/.pw-wrap")
        s, url, login_page = req(publik, "/login")
        must(s == 200 and 'type="password"' in login_page
             and "/static/app.js" in login_page,
             "GET /login -> ada kolom password dan app.js dimuat (toggle mata aktif)")

        s, url, _ = req(publik, "/admin")
        must("/login" in url, "GET /admin tanpa login diarahkan ke /login")

        s, url, page = req(publik, "/login", {"username": "admin", "password": "salah"})
        must("/login" in url and "Username atau password salah" in page,
             "login dengan password salah ditolak")

        admin = client()
        url = login_admin(admin)
        must("/admin" in url, "login admin berhasil")
        s, url, dash = req(admin, "/admin")
        must("masih memakai password awal" in dash,
             "dashboard menampilkan peringatan password awal")

        s, url, page = req(admin, "/admin/pengaturan", PENGATURAN)
        must("Pengaturan disimpan" in page, "pengaturan (nama/jadwal/hasil_live) tersimpan")
        s, url, home = req(publik, "/")
        must("Pemilihan Ketua OSIS 2026" in home and "belum dibuka" in home,
             "halaman pemilih menampilkan nama pemilu dan info belum dibuka")

        # --- kandidat CRUD ---
        s, url, page = req_multipart(
            admin, "/admin/kandidat/baru",
            {"nomor_urut": "1", "nama": "Andi Pratama", "kelas": "XI IPA 1",
             "visi": "Sekolah bersih dan aktif"},
            {"foto": ("andi.png", PNG_1PX, "image/png")})
        must("Andi Pratama" in page, "kandidat 1 (dengan foto) ditambahkan")
        req_multipart(admin, "/admin/kandidat/baru",
                      {"nomor_urut": "2", "nama": "Sari Lestari", "kelas": "XI IPA 2",
                       "visi": "Suara siswa didengar"}, {})
        req_multipart(admin, "/admin/kandidat/baru",
                      {"nomor_urut": "3", "nama": "Rizky Maulana", "kelas": "X IPA 1",
                       "visi": "Ekskul makin hidup"}, {})
        s, url, page = req_multipart(
            admin, "/admin/kandidat/baru",
            {"nomor_urut": "1", "nama": "Duplikat Nomor", "kelas": "X", "visi": ""}, {})
        must("sudah dipakai" in page, "nomor urut duplikat ditolak")
        s, url, daftar = req(admin, "/admin/kandidat")
        kids = re.findall(r'/admin/kandidat/(\d+)"', daftar)
        must(kids[:3] == ["1", "2", "3"], f"tiga kandidat terdaftar (id {kids[:3]})")
        s, url, form1 = req(admin, "/admin/kandidat/1")
        m_foto = re.search(r"/foto/([a-f0-9]{24}\.png)", form1)
        must(m_foto is not None, "foto kandidat tersimpan dengan nama token")
        r = admin.open(BASE + "/foto/" + m_foto.group(1))
        must(r.status == 200 and r.headers.get_content_type() == "image/png"
             and r.read() == PNG_1PX, "foto kandidat bisa diambil kembali lewat /foto/")
        s, url, page = req_multipart(
            admin, "/admin/kandidat/2/edit",
            {"nomor_urut": "2", "nama": "Sari Lestari", "kelas": "XI IPA 2",
             "visi": "Suara siswa didengar dan ditindaklanjuti"}, {})
        must("disimpan" in page, "edit kandidat berhasil")
        s, url, cari = req(admin, "/admin/kandidat?q=Andi")
        must("Andi Pratama" in cari and "Sari Lestari" not in cari,
             "pencarian kandidat menyaring hasil")

        # --- pemilih CRUD + impor ---
        s, url, page = req(admin, "/admin/pemilih/baru",
                           {"nis": "1001", "nama": "Bima", "kelas": "XI IPA 1"})
        must("ditambahkan" in page, "pemilih tunggal ditambahkan")
        s, url, page = req(admin, "/admin/pemilih/baru",
                           {"nis": "1001", "nama": "Bima Lagi", "kelas": "XI IPA 1"})
        must("sudah terdaftar" in page, "NIS duplikat ditolak")
        impor = ("1002;Citra;XI IPA 1\n1003;Dodi;XI IPA 2\n"
                 "1002;Duplikat;XI IPA 1\nbaris rusak\n1004;;XI IPA 2\n"
                 "1005;Eka;XI IPA 2\n1006;Fajar;X IPA 1\n1007;Gita;X IPA 1")
        s, url, page = req(admin, "/admin/pemilih/impor", {"data": impor})
        must("5 pemilih ditambahkan, 3 baris dilewati" in page,
             "impor massal: 5 ditambahkan, 3 dilewati (duplikat/rusak/kosong)")
        s, url, daftar_p = req(admin, "/admin/pemilih")
        must("Belum" in daftar_p and "1007" in daftar_p,
             "daftar pemilih menampilkan NIS dan lencana Belum Memilih")

        # --- voting masih tertutup ---
        s, url, page = req(publik, "/cek-nis", {"nis": "1001"})
        must("belum dibuka" in page, "cek NIS sebelum jam buka -> info belum dibuka")
        s, url, page = req(publik, "/surat-suara")
        must("Verifikasi NIS dulu" in page,
             "surat suara tanpa verifikasi ditolak")
        s, url, page = req(publik, "/suara", {})
        must("Verifikasi NIS dulu" in page, "POST /suara tanpa sesi ditolak")
        s, url, layar = req(publik, "/layar")
        must(s == 200 and "Total Pemilih" in layar and "Partisipasi" in layar,
             "GET /layar publik tanpa login -> 200 berisi agregat")
        must("1001" not in layar and "Bima" not in layar,
             "/layar TIDAK memuat identitas pemilih (NIS/nama)")
    finally:
        berhenti(proc)

    # =====================================================================
    # FASE B — voting berlangsung (sekarang 10:00)
    # =====================================================================
    proc = mulai(tmp, "2026-11-10T10:00:00")
    try:
        admin = client()
        login_admin(admin)

        def memilih(nis, kid, op=None):
            op = op or client()
            s, url, page = req(op, "/cek-nis", {"nis": nis})
            must("/surat-suara" in url, f"NIS {nis} terverifikasi -> surat suara")
            s, url, page = req(op, "/konfirmasi", {"kandidat": str(kid)})
            must("Suara tidak dapat diubah setelah dikirim" in page,
                 f"NIS {nis}: halaman konfirmasi tampil dengan peringatan")
            s, url, page = req(op, "/suara", {})
            must("/sukses" in url and "suaramu sudah tercatat" in page,
                 f"NIS {nis}: suara terkirim -> halaman sukses")
            return op

        v1 = memilih("1001", 1)
        s, url, page = req(v1, "/cek-nis", {"nis": "1001"})
        must("sudah menggunakan hak pilihmu" in page,
             "memilih kedua kali ditolak di cek NIS")
        s, url, page = req(client(), "/cek-nis", {"nis": "9999"})
        must("NIS tidak ditemukan" in page, "NIS tak dikenal -> pesan tidak ditemukan")

        # Dua sesi untuk pemilih yang sama: sesi kedua tidak bisa mengirim lagi.
        opA, opB = client(), client()
        req(opA, "/cek-nis", {"nis": "1002"})
        req(opB, "/cek-nis", {"nis": "1002"})
        req(opA, "/konfirmasi", {"kandidat": "2"})
        s, url, page = req(opA, "/suara", {})
        must("/sukses" in url, "sesi A untuk NIS 1002 berhasil memilih")
        req(opB, "/konfirmasi", {"kandidat": "1"})
        s, url, page = req(opB, "/suara", {})
        must("sudah menggunakan hak pilihmu" in page and "/sukses" not in url,
             "POST /suara kedua (sesi lain) ditolak: satu NIS satu suara")

        memilih("1003", 1)
        memilih("1005", 1)
        memilih("1006", 2)
        # Suara sekarang: kandidat1=3 (1001,1003,1005), kandidat2=2 (1002,1006)

        s, url, dash = req(admin, "/admin")
        must("83,3%" in dash, "dashboard: partisipasi 5/6 = 83,3%")
        s, url, hasil = req(admin, "/admin/hasil")
        must("Andi Pratama" in hasil and "Sari Lestari" in hasil,
             "halaman hasil memuat kandidat")
        s, url, layar = req(client(), "/layar")
        must("60,0%" in layar and "3 suara" in layar,
             "/layar (hasil_live ON) menampilkan perolehan kandidat teratas")

        # --- kerahasiaan suara: skema tabel suara tanpa kolom pemilih ---
        conn = sqlite3.connect(db_path)
        kolom = [r[1] for r in conn.execute("PRAGMA table_info(suara)")]
        jml_suara = conn.execute("SELECT COUNT(*) FROM suara").fetchone()[0]
        conn.close()
        must(kolom == ["id", "kandidat_id", "waktu"],
             f"tabel suara HANYA berisi {kolom} (tanpa referensi pemilih)")
        must(jml_suara == 5, "lima surat suara tersimpan di database")

        # --- integritas sesudah ada suara ---
        s, url, page = req(admin, "/admin/kandidat/1/hapus", {})
        must("TIDAK dapat dihapus" in page, "kandidat bersuara tidak dapat dihapus")
        s, url, page = req(admin, "/admin/kandidat/3/hapus", {})
        must("dihapus" in page, "kandidat tanpa suara dapat dihapus")
        s, url, cari = req(admin, "/admin/pemilih?q=1001")
        m = re.search(r'/admin/pemilih/(\d+)"', cari)
        pid_1001 = m.group(1)
        s, url, page = req(admin, f"/admin/pemilih/{pid_1001}/hapus", {})
        must("TIDAK dapat dihapus" in page, "pemilih yang sudah memilih tidak dapat dihapus")
        s, url, page = req(admin, f"/admin/pemilih/{pid_1001}/edit",
                           {"nis": "9001", "nama": "Bima", "kelas": "XI IPA 1"})
        must("tidak dapat diubah NIS" in page,
             "NIS pemilih yang sudah memilih tidak dapat diubah")
        s, url, page = req(admin, f"/admin/pemilih/{pid_1001}/edit",
                           {"nis": "1001", "nama": "Bima Saputra", "kelas": "XI IPA 1"})
        must("disimpan" in page, "nama pemilih yang sudah memilih masih boleh diedit")
        s, url, cari = req(admin, "/admin/pemilih?q=1007")
        pid_1007 = re.search(r'/admin/pemilih/(\d+)"', cari).group(1)
        s, url, page = req(admin, f"/admin/pemilih/{pid_1007}/hapus", {})
        must("dihapus" in page, "pemilih yang belum memilih dapat dihapus")

        # --- hasil_live OFF menyembunyikan rincian selama voting ---
        mati = dict(PENGATURAN); mati.pop("hasil_live")
        req(admin, "/admin/pengaturan", mati)
        s, url, layar = req(client(), "/layar")
        must("Andi Pratama" in layar and "Partisipasi" in layar,
             "/layar (hasil_live OFF): nama kandidat & partisipasi tetap tampil")
        must("3 suara" not in layar and "60,0%" not in layar,
             "/layar (hasil_live OFF): jumlah suara & persen disembunyikan")
        req(admin, "/admin/pengaturan", PENGATURAN)

        # --- CSV ---
        s, url, hdr, data = req_bytes(admin, "/admin/hasil/hasil.csv")
        teks = data.decode("utf-8-sig")
        must("Andi Pratama" in teks and "TOTAL" in teks and ",3," in teks,
             "CSV hasil akhir memuat kandidat dan total suara")
        s, url, hdr, data = req_bytes(admin, "/admin/hasil/partisipasi.csv")
        teks = data.decode("utf-8-sig")
        must("XI IPA 1" in teks and "TOTAL" in teks,
             "CSV partisipasi memuat kelas dan total")

        # --- sesi pemilih tidak bisa masuk admin ---
        vp = client()
        req(admin, "/admin/pemilih/baru",
            {"nis": "1008", "nama": "Hana", "kelas": "X IPA 1"})
        s, url, page = req(vp, "/cek-nis", {"nis": "1008"})
        must("/surat-suara" in url, "NIS 1008 terverifikasi (sesi pemilih aktif)")
        s, url, page = req(vp, "/admin")
        must("/login" in url, "sesi PEMILIH ditolak dari area admin")

        # --- backup: unduh + berkas asing ditolak ---
        s, url, hdr, zipdata = req_bytes(admin, "/admin/backup/unduh")
        zf = zipfile.ZipFile(io.BytesIO(zipdata))
        names = zf.namelist()
        manifest = json.loads(zf.read("manifest.json"))
        must(manifest.get("app") == "vote-osis" and "vote.db" in names
             and any(n.startswith("uploads/") for n in names),
             "backup ZIP berisi vote.db, uploads, dan manifest vote-osis")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("manifest.json", json.dumps({"app": "aplikasi-lain"}))
            z.writestr("vote.db", b"bukan database")
        s, url, page = req_multipart(admin, "/admin/backup/pulihkan", {},
                                     {"berkas": ("asing.zip", buf.getvalue(),
                                                 "application/zip")})
        must("bukan backup aplikasi PEMILU OSIS" in page,
             "backup milik aplikasi lain ditolak saat pulihkan")
        s, url, daftar = req(admin, "/admin/kandidat")
        must("Andi Pratama" in daftar, "data tidak berubah sesudah penolakan backup")
    finally:
        berhenti(proc)

    # =====================================================================
    # FASE C — voting sudah ditutup (sekarang 17:00)
    # =====================================================================
    proc = mulai(tmp, "2026-11-10T17:00:00")
    try:
        admin = client()
        login_admin(admin)
        s, url, page = req(client(), "/cek-nis", {"nis": "1008"})
        must("sudah ditutup" in page, "cek NIS sesudah tutup -> info sudah ditutup")

        # hasil_live OFF pun hasil tampil penuh sesudah tutup
        mati = dict(PENGATURAN); mati.pop("hasil_live")
        req(admin, "/admin/pengaturan", mati)
        s, url, layar = req(client(), "/layar")
        must("60,0%" in layar,
             "/layar sesudah tutup menampilkan hasil penuh walau hasil_live OFF")
        req(admin, "/admin/pengaturan", PENGATURAN)

        # --- Excel bukti ---
        s, url, hdr, data = req_bytes(admin, "/admin/hasil/hasil.xlsx")
        must(s == 200 and hdr.get_content_type()
             == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             "GET hasil.xlsx -> 200 dengan content-type xlsx")
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data))
        must(wb.sheetnames == ["Hasil Akhir", "Partisipasi Kelas"],
             "workbook berisi sheet Hasil Akhir & Partisipasi Kelas")
        ws = wb["Hasil Akhir"]
        semua = [[c.value for c in row] for row in ws.iter_rows()]
        datar = [str(v) for row in semua for v in row if v is not None]
        must(any("SMA Negeri 1 Contoh" in v for v in datar)
             and any("Pemilihan Ketua OSIS 2026" in v for v in datar),
             "sheet Hasil Akhir memuat nama sekolah & nama pemilu")
        baris_menang = [row for row in semua if "Pemenang" in [str(v) for v in row]]
        must(len(baris_menang) == 1 and "Andi Pratama" in baris_menang[0]
             and 3 in baris_menang[0],
             "baris Pemenang = Andi Pratama dengan 3 suara (tertinggi)")
        must(any(v == "Suara Masuk: 5" for v in datar)
             and any(v == "Total Pemilih Terdaftar: 6" for v in datar),
             "ringkasan Excel cocok dengan data (6 pemilih, 5 suara)")
        must(any(v.startswith("Voting ditutup:") and "10 November 2026 pukul 16.00" in v
                 for v in datar),
             "baris bukti: 'Voting ditutup: 10 November 2026 pukul 16.00 WIB'")
        must(any(v.startswith("Laporan dibuat:") for v in datar),
             "baris bukti: 'Laporan dibuat: ...' tercantum")
        must(ws.protection.sheet is True
             and wb["Partisipasi Kelas"].protection.sheet is True,
             "kedua sheet Excel TERKUNCI (worksheet protection aktif)")
        ws2 = wb["Partisipasi Kelas"]
        datar2 = [str(c.value) for row in ws2.iter_rows() for c in row
                  if c.value is not None]
        must("XI IPA 1" in datar2 and "X IPA 1" in datar2,
             "sheet Partisipasi Kelas memuat kelas-kelas terdaftar")

        # --- Berita Acara ---
        s, url, ba = req(admin, "/admin/berita-acara")
        must(s == 200 and "BERITA ACARA HASIL PEMILIHAN" in ba,
             "Berita Acara tampil dengan judul resmi")
        must("Andi Pratama" in ba and "SMA Negeri 1 Contoh" in ba
             and "10 November 2026 pukul 16.00" in ba,
             "Berita Acara memuat pemenang, sekolah, dan waktu tutup")
        must("Ketua Panitia" in ba and "Kepala Sekolah" in ba,
             "Berita Acara memuat blok tanda tangan Ketua Panitia & Kepala Sekolah")
        must("window.print()" in ba and "Cetak / Simpan PDF" in ba,
             "Berita Acara punya tombol Cetak / Simpan PDF (window.print)")
        must("side-nav" not in ba, "Berita Acara TANPA kerangka admin (siap cetak)")

        # --- ganti password ---
        s, url, page = req(admin, "/ganti-password",
                           {"lama": "admin123", "baru": "rahasia-baru-1",
                            "ulang": "rahasia-baru-1"})
        must("Password berhasil diganti" in page, "ganti password berhasil")
        s, url, dash = req(admin, "/admin")
        must("masih memakai password awal" not in dash,
             "peringatan password awal hilang sesudah diganti")
        req(admin, "/keluar")
        s, url, page = req(admin, "/login", {"username": "admin", "password": "admin123"})
        must("Username atau password salah" in page,
             "password lama tidak berlaku lagi")
        url = login_admin(admin, "rahasia-baru-1")
        must("/admin" in url, "login dengan password baru berhasil")
    finally:
        berhenti(proc)

    # =====================================================================
    # FASE D — pulihkan backup di server BARU yang kosong
    # =====================================================================
    tmp2 = tempfile.mkdtemp(prefix="vote-test-pulih-")
    proc = mulai(tmp2, "2026-11-10T10:00:00")
    try:
        admin2 = client()
        login_admin(admin2)
        s, url, page = req_multipart(admin2, "/admin/backup/pulihkan", {},
                                     {"berkas": ("backup-vote.zip", zipdata,
                                                 "application/zip")})
        must("/login" in url, "pulihkan backup -> sesi dibersihkan, kembali ke login")
        url = login_admin(admin2)
        must("/admin" in url, "login lagi sesudah pulihkan (password dari backup)")
        s, url, daftar = req(admin2, "/admin/kandidat")
        must("Andi Pratama" in daftar and "Sari Lestari" in daftar,
             "kandidat kembali dari backup")
        s, url, daftar_p = req(admin2, "/admin/pemilih")
        must("1001" in daftar_p and "1008" in daftar_p,
             "pemilih kembali dari backup")
        s, url, hasil = req(admin2, "/admin/hasil")
        must("Andi Pratama" in hasil, "hasil suara kembali dari backup")
        conn = sqlite3.connect(os.path.join(tmp2, "sub", "vote.db"))
        jml = conn.execute("SELECT COUNT(*) FROM suara").fetchone()[0]
        conn.close()
        must(jml == 5, "lima suara ikut pulih dari backup")
        s, url, form1 = req(admin2, "/admin/kandidat/1")
        m_foto = re.search(r"/foto/([a-f0-9]{24}\.png)", form1)
        must(m_foto is not None, "foto kandidat ikut pulih dari backup")
        r = admin2.open(BASE + "/foto/" + m_foto.group(1))
        must(r.read() == PNG_1PX, "berkas foto utuh sesudah pulihkan")
    finally:
        berhenti(proc)

    print(f"\nSEMUA UJI LULUS ({JUMLAH[0]} pemeriksaan)")


if __name__ == "__main__":
    main()
