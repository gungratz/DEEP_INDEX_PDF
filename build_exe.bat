@echo off
pyinstaller --noconfirm --windowed --onefile ^
--add-data "SumatraPDF;SumatraPDF" ^
--icon="DeepIndexPDF_icon.ico" ^
--name "DeepIndexPDF" .\pdf_search_gui_full3.py
pause
