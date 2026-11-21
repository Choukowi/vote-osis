# PEMILU OSIS — e-Voting Sekolah

Aplikasi web pemungutan suara elektronik untuk pemilihan Ketua OSIS
(atau pemilihan sekolah sejenis). Dibangun dengan Python 3 standard
library (`http.server` + `sqlite3`), tanpa framework web. Satu-satunya
dependensi eksternal adalah **openpyxl**, khusus untuk ekspor Excel.

## Tiga wajah aplikasi

1. **Panel Admin (gelap, navy/toska/emas)** — dashboard hasil langsung
   (podium, grafik batang, donat partisipasi, suara per kelas, muat ulang
   otomatis 30 detik), kelola kandidat, kelola pemilih (+ impor massal),
   hasil suara, pengaturan, backup & pulihkan, ganti password.
2. **Halaman Pemilih (terang, biru/putih/hijau)** — siswa membuka alamat
   utama `/`, memasukkan NIS, memilih kandidat di surat suara, memeriksa
   halaman konfirmasi, lalu mengirim suara. Selesai: suara tercatat dan
   tidak bisa diubah.
3. **Mode Layar Tampil `/layar` (publik, layar penuh)** — papan hasil
   untuk proyektor/TV: statistik, podium, grafik, partisipasi per kelas.
   Hanya angka agregat — tidak ada identitas pemilih.

## Kerahasiaan suara (penting)

Tabel surat suara (`suara`) hanya menyimpan `kandidat_id` dan waktu —
**tanpa kolom pemilih sama sekali**. Sistem tahu SIAPA yang sudah
memilih (untuk mencegah suara ganda), tetapi tidak pernah bisa tahu
SIAPA memilih siapa. Admin pun hanya melihat partisipasi, bukan pilihan.

## Menjalankan

```bash
cd vote-osis
pip install -r requirements.txt   # hanya openpyxl
python3 app.py                    # http://localhost:8000
```

Variabel lingkungan:

| Variabel | Default | Keterangan |
|---|---|---|
| `PORT` | `8000` | Port server |
| `VOTE_DB` | `data/vote.db` | Lokasi database SQLite |
| `VOTE_UPLOADS` | `data/uploads/` | Folder foto kandidat & logo |
| `VOTE_NOW_OVERRIDE` | — | Override "sekarang" (ISO datetime) — hanya untuk pengujian |

## Akun admin awal

- Username: `admin`
- Password: `admin123`

**Segera ganti password** lewat menu Ganti Password setelah pertama
kali masuk. Selama password masih `admin123`, dashboard menampilkan
peringatan. Semua kolom password dilengkapi tombol mata (lihat/sembunyi).

## Alur pemakaian oleh panitia

1. Masuk sebagai admin → **Pengaturan**: isi nama pemilu, nama sekolah,
   logo, serta **waktu buka** dan **waktu tutup** voting. Tentukan juga
   apakah hasil per kandidat boleh terlihat publik selama voting
   berlangsung (*Hasil Live*).
2. **Kandidat** → tambah calon (nomor urut, nama, kelas, visi, foto).
   Kandidat yang sudah memperoleh suara tidak dapat dihapus.
3. **Pemilih** → tambah siswa satu per satu atau **Impor Massal**:
   satu baris per siswa dengan format `NIS;Nama;Kelas`.
   Pemilih yang sudah memilih tidak dapat dihapus dan NIS-nya terkunci.
4. Bagikan alamat aplikasi ke siswa / perangkat bilik suara. Di luar
   jendela waktu buka–tutup, voting otomatis tertutup.
5. Pantau lewat **Dashboard** atau tampilkan **`/layar`** di proyektor.
6. Sesudah tutup: unduh hasil dari **Hasil Suara** (Excel/CSV) dan/atau
   cetak **Berita Acara Hasil** resmi.

## Unduhan hasil & bukti resmi

- **Excel (.xlsx)** — sheet *Hasil Akhir* (pemenang disorot & diberi label
  "Pemenang", ringkasan partisipasi, baris bukti **"Voting ditutup: …"**
  dan **"Laporan dibuat: …"**) dan sheet *Partisipasi Kelas*.
  **Kedua sheet dikunci** dengan proteksi lembar kerja Excel (kata kunci
  di konstanta `KUNCI_EXCEL` dalam `app.py`) supaya berkas bukti tidak
  mudah disunting. Catatan jujur: proteksi ini adalah **pencegah
  suntingan kasual, bukan enkripsi** — orang yang paham teknis tetap
  bisa membukanya. Untuk bukti yang benar-benar tidak dapat diubah,
  gunakan Berita Acara dalam bentuk PDF.
- **Berita Acara Hasil** (`/admin/berita-acara`) — dokumen resmi siap
  cetak: kop sekolah & nama pemilu, tabel hasil, ringkasan, waktu
  penutupan, dan blok tanda tangan **Ketua Panitia** serta
  **Mengetahui, Kepala Sekolah**. Tombol **Cetak / Simpan PDF**
  (`window.print()`, tombol otomatis hilang saat dicetak) — simpan
  sebagai PDF dari browser; PDF itulah salinan resmi yang tidak bisa
  disunting.
- **CSV** — hasil akhir dan partisipasi per kelas, untuk diolah lebih
  lanjut bila perlu.

## Backup & Pulihkan

Menu **Backup** (admin): unduh satu berkas ZIP berisi database
(`vote.db`), semua foto/logo, dan `manifest.json`. Untuk memulihkan,
unggah ZIP tersebut di halaman yang sama — aplikasi memeriksa manifest
(backup aplikasi lain ditolak), menyimpan salinan pengaman
(`vote-sebelum-pulih.db`), lalu meminta login ulang. **Cadangkan secara
berkala**, terutama sebelum dan sesudah hari pemungutan suara.

## Uji otomatis

```bash
python3 test_alur.py
```

Menjalankan server uji pada database sementara (port 8793, dapat diubah
lewat `TEST_PORT`) dan memeriksa seluruh alur: login, CRUD kandidat &
pemilih, impor, jendela voting, suara ganda ditolak, kerahasiaan skema
suara, hasil live, CSV, Excel terkunci + baris bukti, Berita Acara,
backup–pulihkan, dan penjagaan akses admin. Berakhir dengan
`SEMUA UJI LULUS` bila semuanya beres.

## Berkas

```
app.py            aplikasi utama (server, rute, logika)
static/app.js     toggle mata password + hitung mundur dashboard
static/style.css  tema gelap admin/layar + tema terang pemilih
test_alur.py      uji end-to-end
requirements.txt  openpyxl (satu-satunya dependensi)
Procfile          web: python app.py
runtime.txt       python-3.12.8
data/             database & unggahan (dibuat otomatis, masuk .gitignore)
```
