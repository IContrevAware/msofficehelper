@echo off
setlocal EnableDelayedExpansion
title ODT - Helper

echo =========================================
echo Microsoft Office Deployment Tool Helper
echo       V1.0.1 help plz
echo =========================================
echo.

echo Read instructions on GitHub
pause

set SETUP_EXE=setup.exe
set XMLFILE=office_config.xml

if not exist "%SETUP_EXE%" (
    echo setup.exe not found. Exiting.
    pause
    exit /b
)

echo Select Office Product:
echo 1. Microsoft 365 Apps for Enterprise (Retail)
echo 2. Microsoft 365 Apps for Business (Retail)
echo 3. Microsoft 365 Enterprise No Teams (Retail)
echo 4. Office LTSC / Volume 2024 Pro Plus
echo 5. Office LTSC / Volume 2024 Standard
echo 6. Office 2021 Pro Plus Volume
echo 7. Office 2021 Standard Volume
echo 8. Project Professional 2021 Volume
echo 9. Project Standard 2021 Volume
echo 10. Visio Professional 2021 Volume
echo 11. Visio Standard 2021 Volume

:product_select
set /p PRODNUM=Enter number [1-11]: 

if "%PRODNUM%"=="1" set PRODUCT=O365ProPlusRetail&set CHANNEL=Current
if "%PRODNUM%"=="2" set PRODUCT=O365BusinessRetail&set CHANNEL=Current
if "%PRODNUM%"=="3" set PRODUCT=O365ProPlusEEANoTeamsRetail&set CHANNEL=Current
if "%PRODNUM%"=="4" set PRODUCT=ProPlus2024Volume&set CHANNEL=PerpetualVL2024
if "%PRODNUM%"=="5" set PRODUCT=Standard2024Volume&set CHANNEL=PerpetualVL2024
if "%PRODNUM%"=="6" set PRODUCT=ProPlus2021Volume&set CHANNEL=PerpetualVL2021
if "%PRODNUM%"=="7" set PRODUCT=Standard2021Volume&set CHANNEL=PerpetualVL2021
if "%PRODNUM%"=="8" set PRODUCT=ProjectPro2021Volume&set CHANNEL=PerpetualVL2021
if "%PRODNUM%"=="9" set PRODUCT=ProjectStd2021Volume&set CHANNEL=PerpetualVL2021
if "%PRODNUM%"=="10" set PRODUCT=VisioPro2021Volume&set CHANNEL=PerpetualVL2021
if "%PRODNUM%"=="11" set PRODUCT=VisioStd2021Volume&set CHANNEL=PerpetualVL2021

if not defined PRODUCT (
    echo Invalid selection, try again.
    goto product_select
)

set /p LANGUAGE=Enter language code (e.g., en-gb, en-us): 

echo.
echo Select apps you want to INSTALL (by number, space-separated):
echo 1. Word
echo 2. Excel
echo 3. PowerPoint
echo 4. Outlook
echo 5. OneNote
echo 6. Access
echo 7. Publisher
echo 8. Teams
echo e.g : 1 3 6 (1 3 6 gets word, ppt and access)

set /p APPNUMS=Enter selection: 

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

echo.
set /p INSTALLTYPE=Offline or Online install? (OFFLINE/ONLINE): 

echo Generating configuration XML...

(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="%CHANNEL%"^>
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
