# WebUI Rekap Presensi Pegawai (PDF & Word)

Aplikasi WebUI berbasis Python Flask untuk membuat dan mencetak **Rekap Presensi Pegawai** (SDN Sendangmulyo 04 Semarang) dalam format **PDF** dan **Microsoft Word (.docx)** dengan tata letak kertas A4 yang presisi.

## Fitur Utama
- **Live Preview A4:** Pratinjau dokumen secara langsung saat data diubah.
- **Pemilih Kalender & Waktu:** Input tanggal otomatis dikonversi ke format hari bahasa Indonesia (`SENIN, DD-MM-YYYY`), dilengkapi pemilih waktu hingga satuan detik.
- **Upload Logo Instansi:** Mendukung unggah logo instansi untuk kop surat.
- **Export Langsung:** Unduh hasil akhir langsung dalam format `.pdf` (via WeasyPrint) maupun `.docx` (via python-docx).

## Cara Instalasi & Menjalankan (Linux / WSL)

1. **Clone repository ini:**
   ```bash
   git clone [https://github.com/sinyoabe/rekap-presensi-webui.git](https://github.com/sinyoabe/rekap-presensi-webui.git)
   cd rekap-presensi-webui
