@echo off
if not exist "bin" mkdir "bin"
if exist ".\bin\macro_pack.exe" DEL ".\bin\macro_pack.exe"
MKDIR "build_tmp" 
XCOPY "src" "build_tmp\" /E
COPY "assets\Vintage-Gramophone.ico" "build_tmp"
CHDIR "build_tmp"
pyinstaller --clean --onefile ^
	 -p modules -p pro_modules -p common ^
	 --key=%random%%random%%random% ^
	 --icon "Vintage-Gramophone.ico" ^
	 macro_pack.py
COPY "dist\macro_pack.exe" "..\bin\macro_pack.exe"
CHDIR ..
RMDIR /s /q "build_tmp" 
PAUSE