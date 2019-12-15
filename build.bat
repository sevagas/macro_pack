@echo off
if not exist "bin" mkdir "bin"
if exist ".\bin\macro_pack.exe" DEL ".\bin\macro_pack.exe"
if exist "%TEMP%\build_tmp" RMDIR /s /q "%TEMP%\build_tmp" 
MKDIR "%TEMP%\build_tmp" 
XCOPY "src" "%TEMP%\build_tmp\" /E

COPY "assets\Vintage-Gramophone.ico" "%TEMP%\build_tmp/"
COPY "assets\upx.exe" "%TEMP%\build_tmp/"

set "currentDir=%cd%"
CHDIR /D "%TEMP%\build_tmp\"

pyinstaller --clean --onefile --upx-exclude=vcruntime140.dll -p modules -p common --icon "Vintage-Gramophone.ico" macro_pack.py

CHDIR /D %currentDir%
COPY "%TEMP%\build_tmp\dist\macro_pack.exe" "bin\macro_pack.exe"

RMDIR /s /q "%TEMP%\build_tmp" 
PAUSE