@echo off
REM تشغيل قاعدة بيانات مضرب تعلب: دوس دبل كليك على الملف ده
cd /d "%~dp0"
start "" /min python -m http.server 8000
timeout /t 2 /nobreak >nul
start http://localhost:8000/
echo القاعدة شغالة، اقفل الشباك ده لما تخلص
pause >nul
