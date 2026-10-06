# ORBID Report Manager

Generator laporan ORBID Manager View: HTML statis yang me-render report 5–7 halaman A4 (replika desain Figma) dari data agregat Excel, siap dicetak/disimpan sebagai PDF.

## Cara pakai

```bash
python3 -m http.server 8731
```

Buka `http://localhost:8731`:

1. **Unduh template Excel** — berisi 8+ sheet (`Info`, `Anggota`, `Komitmen`, `Kompetensi`, `Penilai`, `Kontribusi`, `STAR`, `Tren`, `Upward`, `Upward Uraian`, `Insight`) terisi data contoh fiktif sebagai acuan format.
2. Isi dengan data tim (jumlah anggota & nama item/aspek fleksibel, dibaca dari header).
3. Drop file ke halaman → report ter-render.
4. **Cetak / Simpan PDF** — ukuran halaman 798×1122 px, warna ikut tercetak.

## Halaman

| Hal | Isi |
|---|---|
| 1 | Profil 3C tim: kartu Commitment/Competency/Contribution + tren, peta kuadran Contribution×Competency (warna bubble = Commitment), sorotan anggota |
| 2 | Tabel seluruh anggota (3C, tren, blindspot, kategori) + kompetensi tim vs standar |
| 3 | Detail Commitment per item (mis. Gallup Q12) + keterangan item |
| 4 | Detail Competency per aspek + perbandingan perspektif penilai |
| 5 | Detail Contribution per aspek + ringkasan bukti STAR tervalidasi |
| 6 | Upward feedback: average + sebaran nilai per item (opsional, dari sheet `Upward`) |
| 7 | Uraian open feedback untuk manager (opsional, dari sheet `Upward Uraian`) |

Sheet `Tren`, `Upward`, dan `Upward Uraian` opsional — bila kosong, elemen/halaman terkait otomatis disembunyikan.

## Konverter data mentah (`convert_gti.py`)

Mengubah rekap raw score 360° internal menjadi Excel format template, per tim:

```bash
python3 convert_gti.py "<rekap-raw.xlsx>" <handle-leader>
```

Konfigurasi per tim (nama anggota, ringkasan bukti, insight) dibaca dari `team_config.py` — file ini berisi data personal sehingga **di-gitignore**; tanpa file itu konverter tetap jalan dengan nilai default.

> **Catatan**: seluruh file `.xlsx` dan `.pdf` di-gitignore karena berisi data asesmen asli. Jangan commit data karyawan ke repo ini.
