@echo off
setlocal EnableDelayedExpansion
title Office Deployment Tool - User-Friendly Version

rem ==============================================
rem Automated Office Installer with menus
rem ==============================================
echo =========================================
echo Microsoft Office Deployment Tool Installer
echo =========================================
echo.

rem -----------------------------
rem Check for setup.exe
rem -----------------------------
set SETUP_EXE=setup.exe
set XMLFILE=office_config.xml

if not exist "%SETUP_EXE%" (
    echo setup.exe not found. Download the Office Deployment Tool:
    echo https://www.microsoft.com/en-us/download/details.aspx?id=49117
    pause
    exit /b
)

rem -----------------------------
rem Office product menu
rem -----------------------------
echo Select Office Product:
echo 1. Microsoft 365 Apps for business (O365ProPlusRetail)
echo 2. Microsoft Office 2021 Pro Plus Volume (ProPlus2021Volume)
echo 3. Microsoft Office Standard 2021 Volume (Standard2021Volume)

:product_select
set /p PRODNUM=Enter number [1-3]: 
if "%PRODNUM%"=="1" set PRODUCT=O365ProPlusRetail
if "%PRODNUM%"=="2" set PRODUCT=ProPlus2021Volume
if "%PRODNUM%"=="3" set PRODUCT=Standard2021Volume
if not defined PRODUCT (
    echo Invalid selection, try again.
    goto product_select
)

rem -----------------------------
rem Language
rem -----------------------------
set /p LANGUAGE=Enter language code (e.g., en-gb, en-us): 

rem -----------------------------
rem Apps menu
rem -----------------------------
echo.
echo Select apps you want to INSTALL:
echo Type numbers separated by spaces. Example: 1 2 3
echo 1. Word
echo 2. Excel
echo 3. PowerPoint
echo 4. Outlook
echo 5. OneNote
echo 6. Access
echo 7. Publisher
echo 8. Teams

set /p APPNUMS=Enter selection: 

rem Map numbers to apps
set ALL_APPS=Word Excel PowerPoint Outlook OneNote Access Publisher Groove Lync Teams
set INCLUDED_APPS=

for %%A in (%APPNUMS%) do (
    if %%A==1 set INCLUDED_APPS=!INCLUDED_APPS! Word
    if %%A==2 set INCLUDED_APPS=!INCLUDED_APPS! Excel
    if %%A==3 set INCLUDED_APPS=!INCLUDED_APPS! PowerPoint
    if %%A==4 set INCLUDED_APPS=!INCLUDED_APPS! Outlook
    if %%A==5 set INCLUDED_APPS=!INCLUDED_APPS! OneNote
    if %%A==6 set INCLUDED_APPS=!INCLUDED_APPS! Access
    if %%A==7 set INCLUDED_APPS=!INCLUDED_APPS! Publisher
    if %%A==8 set INCLUDED_APPS=!INCLUDED_APPS! Teams
)

rem -----------------------------
rem Offline/Online
rem -----------------------------
echo.
set /p INSTALLTYPE=Offline or Online install? (OFFLINE/ONLINE): 

rem -----------------------------
rem Generate XML
rem -----------------------------
echo Generating configuration XML...

(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="Current"^>
echo     ^<Product ID="%PRODUCT%"^>
echo       ^<Language ID="%LANGUAGE%" /^>

for %%A in (%ALL_APPS%) do (
    set FOUND=0
    for %%I in (%INCLUDED_APPS%) do (
        if /I "%%A"=="%%I" set FOUND=1
    )
    if "!FOUND!"=="0" (
        echo       ^<ExcludeApp ID="%%A" /^>
    )
)

echo     ^</Product^>
echo   ^</Add^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^>
echo ^</Configuration^>
) > %XMLFILE%

echo.
echo Configuration file created: %XMLFILE%
echo.

rem -----------------------------
rem Install logic
rem -----------------------------
if /I "%INSTALLTYPE%"=="OFFLINE" (
    echo Downloading Office files for offline install...
    "%SETUP_EXE%" /download %XMLFILE%
    echo Download finished. Installing...
    "%SETUP_EXE%" /configure %XMLFILE%
) else (
    echo Installing Office (online)...
    "%SETUP_EXE%" /configure %XMLFILE%
)

echo.
echo =====================================
echo Office installation complete.
echo =====================================
pause