# e-Peperiksaan KISLB — v0.5

Tambahan utama v0.5:
- Profil & Keputusan Individu Murid untuk Admin.
- Sejarah semua peperiksaan yang mempunyai markah.
- GPS individu dan purata markah bagi setiap peperiksaan.
- Trend Gerak Gempur individu mengikut siri dengan status Meningkat / Menurun / Kekal.
- Akses terus kepada slip individu.
- Sejarah kelas kekal dalam profil murid.

# e-Peperiksaan KISLB v0.3

Prototaip sistem rekod peperiksaan KAFA Integrasi Sungai Lang Baru.

## Tambahan v0.3
- Pengurusan akaun guru melalui Admin.
- Tugasan guru mengikut kelas + mata pelajaran.
- Pengurusan akaun ibu bapa dan pautan anak.
- Import murid melalui Excel (.xlsx) + template rasmi sistem.
- Naik / tukar kelas secara pukal tanpa memadam sejarah kelas lama.
- Audit log untuk perubahan utama.
- Dashboard Admin yang lebih lengkap.
- Analisa GPS mengikut kelas dan mata pelajaran.
- Graf taburan gred setiap mata pelajaran.
- Trend GPS Gerak Gempur keseluruhan dan setiap mata pelajaran.
- Cetak laporan analisa terus dari browser.

## Skema markah semasa
### PSRA / UASA / UPSA
- Tauhid 50 + Feqah 50 = 100
- Akhlak 50 + Sirah 50 = 100
- Bahasa Arab = 100
- Tajwid = 100
- Jawi 75 + Imlak 15 + Khat 10 = 100
- Mumtaz 90-100, Jayyid Jiddan 75-89, Jayyid 60-74, Maqbul 40-59, Musaadah 0-39
- Mata GPS: 1,2,3,4,5

### UPKK Bertulis /70
Akidah, Ibadah, Adab, Sirah, Jawi & Khat, Bahasa Arab.
- A 63-70, B 42-62, C 28-41, D 0-27
- Mata GPS A=1, B=2, C=3, D=4

### UPKK /100
Al-Quran & Hafazan, PCHI, Amali Solat.
- A 85-100, B 60-84, C 40-59, D 0-39

## Naik taraf daripada v0.2.1
1. Tutup server dengan `CTRL+C`.
2. Backup fail `epeperiksaan.db` jika ia sudah wujud.
3. Salin kandungan v0.3 ke folder projek lama dan benarkan Replace untuk app.py/templates/static/requirements.txt.
4. Jangan padam `epeperiksaan.db` jika mahu kekalkan data sedia ada.
5. Jalankan `python -m pip install -r requirements.txt`.
6. Jalankan `python app.py`.

## Akaun demo
- Admin: admin / admin123
- Guru: guru / guru123
- Ibu bapa: ibu / ibu123

## Nota
Ini masih server pembangunan tempatan. Jangan gunakan debug server ini sebagai production Internet server.

## v0.4
- Logo rasmi KISLB pada header, login, slip dan laporan analisa.
- Slip keputusan dengan format lebih rasmi dan ruang tandatangan.
- Laporan analisa bercetak dengan kepala laporan rasmi.
- Penunjuk Meningkat / Menurun / Kekal untuk trend GPS Gerak Gempur.
- Penambahbaikan paparan telefon dan cetakan.


## v0.5.1
- Slip ibu bapa tidak lagi memaparkan GPS individu / mata GPS.
- PSRA, UASA dan UPSA: paparkan Jumlah Keseluruhan /500 dan Peratus Keseluruhan (jumlah 5 subjek dibahagi 5).
- Pencapaian keseluruhan menggunakan skala STD100.
- UPKK kekal memaparkan markah dan gred mengikut subjek; tiada jumlah/peratus keseluruhan pada slip.
- Susun atur cetakan A4 dipadatkan.
