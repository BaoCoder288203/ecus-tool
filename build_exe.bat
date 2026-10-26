@echo off
chcp 65001 >nul
echo ========================================================
echo   TOOL DONG GOI ECUS HELPER THANH FILE .EXE
echo ========================================================
echo.

echo [1/3] Dang kiem tra va cai dat thu vien can thiet...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [LOI] Khong cai dat duoc thu vien. Vui long kiem tra lai Python!
    pause
    exit /b
)

echo.
echo [2/3] Dang dong goi EcusHelper.exe bang PyInstaller...
pyinstaller --noconsole --onefile --name "EcusHelper" app.py
if %errorlevel% neq 0 (
    echo [LOI] Dong goi that bai!
    pause
    exit /b
)

echo.
echo [3/3] Dang chuan bi goi phan phoi trong thu muc dist/...
copy /Y Template.xlsx dist\Template.xlsx >nul
if not exist dist\input mkdir dist\input
if not exist dist\output mkdir dist\output

echo.
echo ========================================================
echo   THANH CONG! File EcusHelper.exe da nam trong thu muc dist/
echo   Ban chi can copy thu muc dist/ gui cho nguoi khac dung!
echo ========================================================
echo.
pause
