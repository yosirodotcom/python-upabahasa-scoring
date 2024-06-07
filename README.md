# Create Python Env:
python -m venv C:\Users\ASUS\Documents\pythonenv\python_upabahasa
python -m venv G:\myvenv\python_upabahasa

# To activate this environment, use cmd prompt     
C:\Users\ASUS\Documents\pythonenv\python_upabahasa\Scripts\activate.bat
G:\myvenv\python_upabahasa\Scripts\activate.bat

# To deactivate an active environment, use

# Install library
pip install pandas
pip install tk
pip install customtkinter
pip install openpyxl
pip install missingno
pip install pyinstaller

# Export to exe (create from cmd terminal)
pyinstaller  main.py

# Panduan

1. hasilkan output.xlxs dari main.py kemudian simpan file tersebut pada folder Database Google Drive
2. Update data skor melalui Aplikasi Database UPT Bahasa
3. Pada qry_ept tabel telah di filter pada ISO terakhir kegiatan yang berlangsung, 
   jadi kalo mau di filter ISO yang mau di update, lakukan di sini
4. Buka file ept_all.xlsx kemudian Refresh All
5. Buka file ept_data.ipynb
6. Jika hanya ingin menambah data baru, maka gunakan fungsi appendGsheet. Perhatikan ISO yang mau ditambahkan datanya.
   Jangan sampai ISO yang telah ada
7. Jika hanya ingin update data tabel, maka gunakan fungsi updateGsheet. Perhatikan antara sumber tabel untuk mengupdate 
   disesuaikan dengan nomor row di gsheet