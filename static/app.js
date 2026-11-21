// PEMILU OSIS — e-Voting Sekolah: skrip sisi klien.
// Dua blok yang BERDIRI SENDIRI:
// 1) Toggle "mata" untuk SEMUA kolom password (pola yang sama dengan
//    aplikasi Absen): bungkus tiap input[type=password] dengan .pw-wrap
//    dan tambahkan tombol mata. Klik untuk melihat isi password (type
//    menjadi "text") atau menyembunyikannya lagi (kembali "password"),
//    supaya pengguna bisa mengecek salah ketik. Tombol selalu
//    type="button" agar tidak pernah mengirim formulir.
// 2) Hitungan mundur: elemen .hitung[data-target] (epoch milidetik)
//    diperbarui tiap detik menjadi "HH:MM:SS" (didahului "N hari "
//    bila >= 1 hari). Server sudah merender teks awal, jadi halaman
//    tetap benar walau JavaScript mati.

// ---------------------------------------------------------------------------
// 1) Toggle mata password
// ---------------------------------------------------------------------------
(function () {
  var IKON_MATA = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" '
    + 'stroke="currentColor" stroke-width="1.8" stroke-linecap="round" '
    + 'stroke-linejoin="round" aria-hidden="true">'
    + '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/>'
    + '<circle cx="12" cy="12" r="3"/></svg>';
  var IKON_MATA_CORET = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" '
    + 'stroke="currentColor" stroke-width="1.8" stroke-linecap="round" '
    + 'stroke-linejoin="round" aria-hidden="true">'
    + '<path d="M17.94 17.94A10.61 10.61 0 0 1 12 19c-6.5 0-10-7-10-7a17.63 '
    + '17.63 0 0 1 4.06-4.94M9.9 4.24A9.12 9.12 0 0 1 12 5c6.5 0 10 7 10 7a17.65 '
    + '17.65 0 0 1-2.16 3.19M14.12 14.12a3 3 0 1 1-4.24-4.24"/>'
    + '<path d="M2 2l20 20"/></svg>';

  function pasangToggle(input) {
    if (input.getAttribute('data-pw-toggle')) return;
    input.setAttribute('data-pw-toggle', '1');
    var bungkus = document.createElement('span');
    bungkus.className = 'pw-wrap';
    input.parentNode.insertBefore(bungkus, input);
    bungkus.appendChild(input);
    var tombol = document.createElement('button');
    tombol.type = 'button';
    tombol.className = 'pw-toggle';
    tombol.setAttribute('aria-label', 'Tampilkan password');
    tombol.innerHTML = IKON_MATA;
    bungkus.appendChild(tombol);
    tombol.addEventListener('click', function () {
      var terlihat = input.type === 'text';
      input.type = terlihat ? 'password' : 'text';
      tombol.innerHTML = terlihat ? IKON_MATA : IKON_MATA_CORET;
      tombol.setAttribute('aria-label',
        terlihat ? 'Tampilkan password' : 'Sembunyikan password');
      input.focus();
    });
  }

  function pasangSemua() {
    var kolom = document.querySelectorAll('input[type="password"]');
    for (var i = 0; i < kolom.length; i++) pasangToggle(kolom[i]);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', pasangSemua);
  } else {
    pasangSemua();
  }
})();

// ---------------------------------------------------------------------------
// 2) Hitungan mundur sisa waktu voting
// ---------------------------------------------------------------------------
(function () {
  var els = document.querySelectorAll('.hitung[data-target]');
  if (!els.length) return;

  function dua(n) { return (n < 10 ? '0' : '') + n; }

  function perbarui() {
    var kini = Date.now();
    for (var i = 0; i < els.length; i++) {
      var target = parseInt(els[i].getAttribute('data-target'), 10);
      var sisa = Math.max(0, Math.floor((target - kini) / 1000));
      var hari = Math.floor(sisa / 86400);
      var jam = Math.floor((sisa % 86400) / 3600);
      var menit = Math.floor((sisa % 3600) / 60);
      var detik = sisa % 60;
      els[i].textContent = (hari ? hari + ' hari ' : '')
        + dua(jam) + ':' + dua(menit) + ':' + dua(detik);
    }
  }

  perbarui();
  setInterval(perbarui, 1000);
})();
