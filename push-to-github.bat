@echo off
echo ===================================================
echo   PUSH EPS 08 KE GITHUB: eps08-ai-presentation-generator
echo   Repository: https://github.com/pangestuaiseries/eps08-ai-presentation-generator.git
echo ===================================================
echo.

cd /d "%~dp0"

echo [1/4] Menginisialisasi Git...
if not exist ".git" (
    git init
    git branch -M main
    git remote add origin https://github.com/pangestuaiseries/eps08-ai-presentation-generator.git
)

echo [2/4] Menambahkan seluruh file...
git add .

echo [3/4] Melakukan commit...
git commit -m "feat: update enterprise UI/UX, dark/light theme, and slide deck" || echo Tidak ada commit baru, melanjutkan push...

echo [4/4] Mendorong ke GitHub (main)...
git push -u origin main

echo.
echo ===================================================
echo   SELESAI! Repositori berhasil diperbarui.
echo   Kunjungi: https://github.com/pangestuaiseries/eps08-ai-presentation-generator
echo ===================================================
pause

