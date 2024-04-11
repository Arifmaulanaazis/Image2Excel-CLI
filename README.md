---

# Image2Excel CLI

![Python](https://img.shields.io/badge/python-v3.7+-blue.svg)
![License](https://img.shields.io/badge/license-GPLv3-blue.svg)

Image2Excel CLI adalah sebuah alat baris perintah (CLI) yang memungkinkan Anda mengubah gambar menjadi file Excel, dengan setiap sel Excel mewakili satu piksel gambar, mewarnai setiap sel sesuai dengan warna piksel yang sesuai.

## Dependensi

Pastikan Anda telah menginstal dependensi berikut sebelum menggunakan Image2Excel CLI:

- [argparse](https://docs.python.org/3/library/argparse.html): Untuk mengelola argumen baris perintah.
- [concurrent.futures](https://docs.python.org/3/library/concurrent.futures.html): Untuk pemrosesan paralel.
- [matplotlib](https://matplotlib.org/): Untuk konversi warna.
- [numpy](https://numpy.org/): Untuk manipulasi array.
- [openpyxl](https://openpyxl.readthedocs.io/): Untuk membuat dan memodifikasi file Excel.
- [Pillow](https://python-pillow.org/): Untuk memanipulasi gambar.
- [tqdm](https://github.com/tqdm/tqdm): Untuk menampilkan kemajuan proses.

Anda dapat menginstal semua dependensi dengan menjalankan perintah berikut:

```
pip install matplotlib numpy openpyxl Pillow tqdm
```

## Cara Menggunakan

Anda dapat menggunakan Image2Excel CLI dengan mengikuti langkah-langkah berikut:

1. Pastikan Anda memiliki Python versi 3.7 atau yang lebih baru terpasang di sistem Anda.

2. Jalankan script `Image2Excel.py` dengan menggunakan perintah berikut:

   ```
   python Image2Excel.py --input <path_input_gambar> --downscale <rasio_downscale> --output <path_output_file_excel>
   ```

   - `--input` atau `-i`: Path ke file gambar yang ingin Anda konversi.
   - `--downscale` atau `-d`: Rasio penurunan skala ukuran gambar (default: 10).
   - `--output` atau `-o`: Path ke file Excel keluaran.

3. Tunggu hingga proses selesai, dan Anda akan memiliki file Excel yang dihasilkan sesuai dengan gambar yang diberikan.

## Contoh Penggunaan

Di bawah ini adalah contoh penggunaan Image2Excel CLI:

```
python Image2Excel.py --input shiro.png --downscale 8 --output output.xlsx
```

## Lisensi

Projek ini dilisensikan di bawah [GPLv3](https://www.gnu.org/licenses/gpl-3.0.html).

## Penulis

Arif Maulana - [GitHub](https://github.com/Arifmaulanaazis)
