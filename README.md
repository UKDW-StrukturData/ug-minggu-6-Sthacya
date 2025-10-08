[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/vVPr144_)
-----

# UG TABULAR: Sistem Pendataan Mahasiswa

Proyek ini adalah tugas untuk membuat sistem pendataan mahasiswa sederhana menggunakan **Streamlit**. Anda berperan sebagai mahasiswa UKDW yang ditugaskan untuk mengimplementasikan fungsionalitas dasar CRUD (Create, Read, Update, Delete) dan pengurutan data yang bersumber dari file Excel.

## 🚀 Fitur

  - **➕ Insert Data**: Menambahkan data mahasiswa baru dengan validasi NIM duplikat.
  - **✏️ Edit Data**: Mengubah data mahasiswa yang sudah ada berdasarkan NIM.
  - **❌ Delete Data**: Menghapus data mahasiswa berdasarkan NIM.
  - **📊 Sort Data**: Mengurutkan tabel data secara dinamis berdasarkan kolom yang dipilih.
  - **🖥️ Interface**: Tampilan antarmuka pengguna yang interaktif dibangun dengan Streamlit.

-----

## 🛠️ Cara Menjalankan Aplikasi

1.  **Pastikan Anda memiliki Python dan library yang dibutuhkan.** Jika belum, install library yang diperlukan:

    ```bash
    pip install streamlit pandas openpyxl
    ```

2.  **Buka terminal atau command prompt** pada direktori utama proyek (folder yang berisi file `app.py`).

3.  **Jalankan perintah berikut:**

    ```bash
    streamlit run app.py
    ```

    Setelah perintah dieksekusi, aplikasi akan otomatis terbuka di browser Anda.

-----

## 📝 Detail Tugas

Anda diminta untuk melengkapi fungsi `Insert`, `Edit`, `Delete`, dan `Sort` pada file `app.py`. Komponen antarmuka seperti *input field*, *selectbox*, dan *button* sudah disediakan.

### ➕ Fungsi `Insert`

  - NIM tidak boleh duplikat.
  - Jika NIM yang diinput sudah ada di dalam file Excel, tampilkan pesan: **"Nim Sudah Ada"**.
  - Jika data berhasil ditambahkan, tampilkan pesan: **"Data Sukses di Masukan"**.
  - Nama tidak boleh angka.

### ✏️ Fungsi `Edit`

  - Data yang akan diubah dipilih berdasarkan **NIM**.
  - Jika NIM yang dituju tidak ada di dalam file Excel, tampilkan pesan: **"Nim tidak ditemukan"**.
  - Data akan diubah sesuai dengan nilai yang diinput pada field `New Nim` dan `New Name`.
  - Jika data berhasil diubah, tampilkan pesan: **"Data Sukses di Edit"**.
  - Nama tidak boleh angka

### ❌ Fungsi `Delete`

  - Data yang akan dihapus dipilih berdasarkan **NIM**.
  - Jika NIM yang dituju tidak ada di dalam file Excel, tampilkan pesan: **"Nim tidak ditemukan"**.
  - Jika data berhasil dihapus, tampilkan pesan: **"Data Sukses di Hapus"**.

### 📊 `Sistem Filter Table`

  - Tabel data yang ditampilkan pada antarmuka Streamlit harus dapat difilter berdasarkan kolom yang dipilih melalui `selectbox`.

-----

## 📜 Aturan Pengerjaan

Mohon patuhi aturan berikut selama pengerjaan:

1.  **JANGAN** mengganti nama `class` atau fungsi yang sudah ada.
2.  **JANGAN** menghapus fungsi yang sudah ada.
3.  **JANGAN** menghapus atau menambah parameter pada `constructor` atau fungsi.
4.  Mengganti nama parameter **DIPERBOLEHKAN**.
5.  **Peringatan**: Aturan di atas boleh dilanggar hanya jika Anda memahami sepenuhnya konsekuensinya dan wajib bisa menjelaskannya.

-----

## 💯 Kriteria Penilaian

  - **Fungsi Insert berjalan baik**: `25%`
  - **Fungsi Edit berjalan baik**: `30%`
  - **Fungsi Delete berjalan baik**: `25%`
  - **Sistem Filter Table berjalan baik**: `10%`
  - **Aplikasi dapat digunakan dengan Streamlit**: `10%`

**Good luck\! ✨**
