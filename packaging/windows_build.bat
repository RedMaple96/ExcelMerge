@echo off
REM Windows 打包脚本：生成 .exe
REM 用法: 双击运行或 cmd 执行 packaging\windows_build.bat
cd /d "%~dp0\.."

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

echo ==^> 开始 PyInstaller 打包...
python -m PyInstaller ^
    --noconfirm ^
    --windowed ^
    --onefile ^
    --name "Excel Merge" ^
    --add-data "src;src" ^
    --hidden-import openpyxl ^
    main.py

echo.
echo ✅ .exe 构建完成: dist\Excel Merge.exe
pause
