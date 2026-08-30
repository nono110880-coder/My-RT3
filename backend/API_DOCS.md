# Dokumentasi API - Sistem Informasi Keuangan RT

Base URL: 'http://10.139.11.19:5005/api'

## Autentikasi (`/auth`)
- `POST /auth/login`
  - **Body:** `{ "username": "...", "password": "..." }`
  - **Response:** Token JWT dan data User.
- `GET /auth/me`
  - **Headers:** `Authorization: Bearer <token>`
  - **Response:** Detail User yang sedang login.

## Warga (`/warga`)
- `GET /warga` : Mengambil seluruh data warga.
- `GET /warga/:id` : Mengambil detail spesifik warga.
- `POST /warga` : Menambahkan data warga (Hanya Admin).
- `PUT /warga/:id` : Memperbarui data warga (Hanya Admin).
- `DELETE /warga/:id` : Menghapus data warga (Hanya Admin).

## Pemasukan (`/pemasukan`)
- `GET /pemasukan` : Daftar histori pemasukan.
- `POST /pemasukan` : Input pemasukan kas RT (Hanya Admin/Bendahara).
- `PUT /pemasukan/:id` : Edit pemasukan (Hanya Admin/Bendahara).
- `DELETE /pemasukan/:id` : Hapus pemasukan (Hanya Admin/Bendahara).

### Kategori Pemasukan (`/pemasukan/kategori`)
- `GET /pemasukan/kategori` : Daftar kategori.
- `POST /pemasukan/kategori` : Tambah kategori.

## Pengeluaran (`/pengeluaran`)
- `GET /pengeluaran` : Daftar histori pengeluaran.
- `POST /pengeluaran` : Input pengeluaran (Hanya Admin/Bendahara).
- `PUT /pengeluaran/:id` : Edit pengeluaran (Hanya Admin/Bendahara).
- `DELETE /pengeluaran/:id` : Hapus pengeluaran (Hanya Admin/Bendahara).

### Kategori Pengeluaran (`/pengeluaran/kategori`)
- `GET /pengeluaran/kategori` : Daftar kategori.
- `POST /pengeluaran/kategori` : Tambah kategori.

## Iuran Warga (`/iuran`)
- `GET /iuran` : Menampilkan seluruh rekap tagihan dan tunggakan warga.
- `POST /iuran/generate`
  - **Fungsi:** Membuat tagihan otomatis untuk bulan & tahun tertentu ke semua warga aktif.
  - **Body:** `{ "bulan": 5, "tahun": 2026, "nominal": 50000 }`
- `POST /iuran/:id/bayar`
  - **Fungsi:** Mengubah status tagihan menjadi `LUNAS` dan otomatis mencatatkan ke `Pemasukan`.

## Dashboard (`/dashboard`)
- `GET /dashboard` : Mengambil metrik (total pemasukan, pengeluaran, saldo saat ini, statistik tunggakan).

## Laporan (`/laporan`)
- `GET /laporan/excel` : Ekspor Buku Kas Umum format `.xlsx` (Otomatis unduh).
- `GET /laporan/pdf` : Ekspor Buku Kas Umum format `.pdf` (Otomatis unduh).
