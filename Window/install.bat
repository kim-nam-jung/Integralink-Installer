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
echo        Integralink Word Add-in Install Program
echo.
echo   ========================================================
echo.

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul 2>&1
if %errorlevel% equ 0 (
    echo   [!] Microsoft Word is running.
    echo.
    echo   Please save your work and close Word completely.
    echo.
    pause
    exit /b
)

set "ADDIN_DIR=C:\OfficeAddin"
set "SHARE_NAME=OfficeAddin"
set "SCRIPT_DIR=%~dp0"

echo   --------------------------------------------------------
echo.
echo   [1/3] Creating install folder...
if not exist "%ADDIN_DIR%" (
    mkdir "%ADDIN_DIR%"
    echo         C:\OfficeAddin folder created.
) else (
    echo         C:\OfficeAddin folder already exists.
)
echo.
echo   [2/3] Copying manifest file...
if exist "%SCRIPT_DIR%manifest.xml" (
    copy /Y "%SCRIPT_DIR%manifest.xml" "%ADDIN_DIR%\manifest.xml" >nul
    echo         manifest.xml copied!
) else (
    echo   [ERROR] manifest.xml file not found.
    pause
    exit /b
)
echo.
echo   [3/3] Setting up network share and catalog...
net share %SHARE_NAME% /delete >nul 2>&1
net share %SHARE_NAME%="%ADDIN_DIR%" /GRANT:Everyone,READ >nul 2>&1

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5F156461-1234-4567-890A-BCDEF0123456}" /v "Id" /t REG_SZ /d "{5F156461-1234-4567-890A-BCDEF0123456}" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5F156461-1234-4567-890A-BCDEF0123456}" /v "Url" /t REG_SZ /d "\\localhost\OfficeAddin" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{5F156461-1234-4567-890A-BCDEF0123456}" /v "Flags" /t REG_DWORD /d 1 /f >nul 2>&1
echo         Network share and catalog configured!
echo.
echo   --------------------------------------------------------
echo.
echo   ========================================================
echo            Installation Complete!
echo   ========================================================
echo.
echo   Please open Word and follow these steps:
echo   1. Go to 'Insert' tab > 'Get Add-ins' (or 'Add-ins')
echo   2. Click 'SHARED FOLDER' tab
echo   3. Select 'Integralink' and click 'Add'
echo.
echo   --------------------------------------------------------
echo.
echo   Press any key to exit...
pause >nul
