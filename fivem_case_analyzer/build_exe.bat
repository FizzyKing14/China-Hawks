@echo off
chcp 65001 >nul
echo ============================================
echo   圣安地列斯州 案件罪名分析器  —  打包 EXE
echo ============================================
echo.
echo 正在安装打包工具(首次较慢)...
python -m pip install --upgrade pyinstaller openpyxl
echo.
echo 正在打包...
pyinstaller --onefile --noconsole --icon "icon.ico" --add-data "icon.ico;." --name "案件罪名分析器" app.py
echo.
echo ============================================
echo   完成！exe 在 dist\ 文件夹里，双击即可用。
echo ============================================
pause
