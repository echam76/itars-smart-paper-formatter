# ITARS Smart Paper Formatter — ITARS 6.0

## Apa yang tersedia
Aplikasi Streamlit memproses `.docx` sehingga 20 MB, mengekalkan struktur dokumen asal sedaya mungkin (termasuk jadual boleh diedit dan objek Word), menetapkan A4, margin 2.54 cm dan Times New Roman, mengemaskini header ITARS 5.0→ITARS 6.0, menyediakan semakan asas dan muat turun Word.

**Penting:** Sistem tidak menulis semula artikel, tidak mengesahkan APA 7, tidak memastikan semua rajah/persamaan terformat sempurna, dan tidak memeriksa kesahihan rujukan. Tahun header diset kepada **2027 sebagai andaian**; ubah nilai `year='2027'` dalam `app.py` apabila tahun ITARS 6.0 disahkan.

## Jalankan secara tempatan
1. Pasang Python 3.10 atau versi lebih baharu.
2. Buka terminal di folder ini, jalankan `python -m pip install -r requirements.txt`.
3. Jalankan `streamlit run app.py`.
4. Buka alamat yang ditunjukkan, lazimnya `http://localhost:8501`.

## Terbitkan ke Streamlit Community Cloud
1. Cipta repositori GitHub dan muat naik fail projek ini (termasuk folder `.streamlit` dan template).
2. Pergi ke https://share.streamlit.io/ dan log masuk dengan GitHub.
3. Pilih repositori, branch dan fail utama `app.py`, kemudian deploy.
4. Semak pautan HTTPS yang dihasilkan, uji muat naik/muat turun dan teliti polisi penggunaan perkhidmatan sebelum menerima manuskrip awam.
5. Streamlit Community Cloud sesuai untuk ujian; penggunaan rasmi berskala besar memerlukan penilaian hosting, kapasiti, privasi dan akses.

## Masukkan ke Google Sites (ITARS 6.0)
1. Dalam Google Sites, cipta halaman `Smart Paper Formatter`.
2. Pilih **Insert → Embed → By URL** dan tampal URL HTTPS aplikasi Streamlit (boleh cuba tambahan `?embed=true`).
3. Besarkan frame dan uji muat naik/muat turun pada komputer dan telefon.
4. Jika embed disekat atau kurang selesa, gunakan pautan/butang `Open Formatter` untuk membuka aplikasi dalam tab baharu.
5. Terbitkan Google Sites selepas ujian sebenar.

## Keselamatan untuk penggunaan rasmi
- Fail diproses dalam memori; kod aplikasi tidak sengaja menulis manuskrip kepada folder atau pangkalan data, tetapi tidak menjamin tiada log/retensi oleh penyedia hosting.
- Semak polisi privasi, hak cipta manuskrip, konfigurasi logs, patch keselamatan, had trafik, antivirus/scan fail dan proses penghapusan fail.
- Tiada authentication/pentadbir/dashboards disertakan dalam prototaip ini.

## Ujian contoh
Manuskrip GBIverse disediakan oleh penganjur untuk ujian; masukkan sebagai `.docx`, jana output dan semak halaman, rajah, jadual, abstrak dan APA 7 secara manual.
