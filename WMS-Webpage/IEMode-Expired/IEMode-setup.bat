@echo off
echo ========================
echo Create the IEMode folder
echo ========================

:: Create IEMode Folder
mkdir "C:\IEMode\" 2>nul

:: Create xml file
echo ========================
echo    Create xml file
echo ========================

set "folder=C:\IEMode\"
set "xmlfile=%folder%sites.xml"
set "regfile=%folder%iemode.reg"

(
echo ^<site-list version="205"^>
echo   ^<created-by^>
echo     ^<tool^>EMIESiteListManager^</tool^>
echo     ^<version^>12.0.0.0^</version^>
echo     ^<date-created^>09/15/2026^</date-created^>
echo   ^</created-by^>
echo   ^<site url="http://Paste URL here"^>
echo     ^<open-in^>IE11^</open-in^>
echo   ^</site^>
echo ^</site-list^>
) > "%xmlfile%"

:: Create regedit file
echo ========================
echo   Create regedit file
echo ========================

(
echo Windows Registry Editor Version 5.00
echo.
echo [HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge]
echo "InternetExplorerIntegrationLevel"=dword:00000001
echo "InternetExplorerIntegrationSiteList"="%xmlfile:\=\\%"
) > "%regfile%"

:: Import the regedit file
echo ========================
echo  Apply regedit settings
echo ========================
reg import "%regfile%"

echo.
echo ========================
echo         All done!
echo ========================
echo XML file  : %xmlfile%
echo REG file  : %regfile%
echo.
echo Close Edge completely, then reopen and check edge://compat/enterprise
pause