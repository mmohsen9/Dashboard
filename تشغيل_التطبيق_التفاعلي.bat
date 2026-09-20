@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo    تشغيل تطبيق الداشبورد التفاعلي - الكوب الاول
echo ============================================================
echo.
echo [1/2] تثبيت/تحديث المكتبات المطلوبة ...
python -m pip install --quiet --upgrade streamlit pandas openpyxl
echo.
echo [2/2] فتح التطبيق في المتصفح ...
python -m streamlit run "تطبيق_الداشبورد_التفاعلي.py"
echo.
pause
