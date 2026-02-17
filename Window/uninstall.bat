@echo off
setlocal EnableDelayedExpansion

net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cls
echo.
echo   ========================================================
echo.
echo        Integralink Word Add-in Uninstall Program
echo.
echo   ========================================================
echo.

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [!] Microsoft Word is running.
    echo.
    echo   Please close Word first.
    echo.
    pause
    exit /b
)

echo   --------------------------------------------------------
echo.
echo   [1/4] Removing startup template...
if exist "%APPDATA%\Microsoft\Word\STARTUP\Integralink.dotm" (
    del /f "%APPDATA%\Microsoft\Word\STARTUP\Integralink.dotm"
    echo         Startup template removed.
) else (
    echo         No startup template found.
)
echo.
echo   [2/4] Removing network share...
net share OfficeAddin /delete >nul 2>&1
echo         Network share removed.
echo.
echo   [3/4] Removing registry entries...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5F156461-1234-4567-890A-BCDEF0123456}" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\integralink" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /f >nul 2>&1
echo         Registry entries removed.
echo.
echo   [4/4] Removing install folder...
if exist "C:\OfficeAddin" (
    rmdir /s /q "C:\OfficeAddin"
    echo         C:\OfficeAddin folder deleted.
) else (
    echo         C:\OfficeAddin folder already deleted.
)
echo.
echo   --------------------------------------------------------
echo.
echo   ========================================================
echo            Uninstall Complete!
echo   ========================================================
echo.
echo   Thank you for using Integralink.
echo.
echo   --------------------------------------------------------
echo.
echo   Press any key to exit...
pause >nul
