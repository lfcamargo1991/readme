@echo off

set "OfficePath=C:\Program Files\Microsoft Office"
set "AltOfficePath=D:\Program Files\Microsoft Office"

IF EXIST "%OfficePath%\PROPLUS\setup.exe" (
    "%OfficePath%\setup.exe" /uninstall PROPLUS /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-proplus.xml
) ELSE IF EXIST "%OfficePath%\STANDARD\setup.exe" (
    "%OfficePath%\setup.exe" /uninstall STANDARD /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-std.xml
) ELSE IF EXIST "%OfficePath%\HOMEBUSINESS\setup.exe" (
    "%OfficePath%\setup.exe" /uninstall HOMEBUSINESS /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-heb.xml
) ELSE IF EXIST "%AltOfficePath%\PROPLUS\setup.exe" (
    "%AltOfficePath%\setup.exe" /uninstall PROPLUS /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-proplus.xml
) ELSE IF EXIST "%AltOfficePath%\HOMEBUSINESS\setup.exe" (
    "%AltOfficePath%\setup.exe" /uninstall HOMEBUSINESS /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-heb.xml
) ELSE IF EXIST "%AltOfficePath%\STANDARD\setup.exe" (
    "%AltOfficePath%\setup.exe" /uninstall STANDARD /dll OSETUP.DLL /config \\ems.com.br\NETLOGON\uninstalloffice2010\config-std.xml
) 

wmic product where "name like 'Microsoft Office Standard 2010%%'" get Name | findstr /C:"Microsoft Office Standard 2010" > nul
if %errorlevel% equ 0 (
    wmic product where "name like 'Microsoft Office Standard 2010%%'" call uninstall /nointeractive

)

wmic product where "name like 'Microsoft Office Professional Plus 2010%%'" get Name | findstr /C:"Microsoft Office Professional Plus 2010" > nul
if %errorlevel% equ 0 (
    wmic product where "name like 'Microsoft Office Professional Plus 2010%%'" call uninstall /nointeractive

)

wmic product where "name like 'Microsoft Office Home and Business 2010%%'" get Name | findstr /C:"Microsoft Office Home and Business 2010" > nul
if %errorlevel% equ 0 (
    wmic product where "name like 'Microsoft Office Home and Business 2010%%'" call uninstall /nointeractive

)


timeout /t 10 >nul

for /f "tokens=2 delims==" %%A in ('wmic product where "name='Microsoft Office Standard 2010'" get IdentifyingNumber /value ^| find "{"') do set ProductID=%%A

if defined ProductID (
    msiexec /x %ProductID% /qn
) else (
    echo O Microsoft Office Standard 2010 não está instalado.
)

timeout /t 10 >nul

for /f "tokens=2 delims==" %%A in ('wmic product where "name='Microsoft Office Professional Plus 2010'" get IdentifyingNumber /value ^| find "{"') do set ProductID=%%A

if defined ProductID (
    msiexec /x %ProductID% /qn
) else (
    echo O Microsoft Office Professional Plus 2010 não está instalado.
)

timeout /t 10 >nul

for /f "tokens=2 delims==" %%A in ('wmic product where "name='Microsoft Office Home and Business 2010'" get IdentifyingNumber /value ^| find "{"') do set ProductID=%%A

if defined ProductID (
    msiexec /x %ProductID% /qn
) else (
    echo O Microsoft Office Home and Bussiness 2010 não está instalado.
)

rem iniciar script powershell
rem powershell.exe -ExecutionPolicy Bypass -File "C:\CAMINHO_DO_SCRIPT_PS\removerOffice2010-MANUAL.ps1"
