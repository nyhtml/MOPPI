@ECHO OFF
COLOR 1F && CLS
TITLE MOPPI v2025.07

:: Get current year using PowerShell (works on all modern systems)
FOR /F %%i IN ('powershell -NoProfile -Command "Get-Date -Format yyyy"') DO SET Year=%%i

:SELECT
CLS
ECHO ---------------------------------------------------------------
ECHO    M.O.P.P.I. - Microsoft Office Professional Plus Installer
ECHO ---------------------------------------------------------------
ECHO.
ECHO    Please make sure that you are connected to the Internet.
ECHO.
ECHO     Looking for older instances of Office in the Control Panel
ECHO               If found, the install will fail.
ECHO.
ECHO ---------------------------------------------------------------
ECHO    Copyright (c) 2013-%Year% Stephan Pringle. All rights reserved.
ECHO                Licensed to Essex County College
ECHO ---------------------------------------------------------------
ECHO.
ECHO  Install:
ECHO  1. Microsoft Office Professional Plus 2019
ECHO  2. Microsoft Office Professional Plus 2019 ^+ Projects ^& Visio
ECHO  3. Microsoft Office LTSC Professional Plus 2021
ECHO  4. Microsoft Office LTSC Professional Plus 2021 ^+ Projects ^& Visio
ECHO  5. Microsoft Office LTSC Professional Plus 2024
ECHO  6. Microsoft Office LTSC Professional Plus 2024 ^+ Projects ^& Visio
ECHO  7. Microsoft Office 365
ECHO.
ECHO 10. Office Removal Tool (The system will reboot)
ECHO.
SET /P ChoosedLanguage=Enter a number and press ENTER key or 0 to quit: 

:: Validate input
SETLOCAL ENABLEDELAYEDEXPANSION
SET /A Dummy=%ChoosedLanguage% >NUL 2>&1
IF ERRORLEVEL 1 (
    ENDLOCAL
    ECHO Invalid input. Please enter a number.
    TIMEOUT /T 2 >NUL
    GOTO SELECT
)
ENDLOCAL

:: Routing
IF "%ChoosedLanguage%"=="0" GOTO ExitApp
IF "%ChoosedLanguage%"=="1" GOTO Install2019
IF "%ChoosedLanguage%"=="2" GOTO Install2019Plus
IF "%ChoosedLanguage%"=="3" GOTO Install2021
IF "%ChoosedLanguage%"=="4" GOTO Install2021Plus
IF "%ChoosedLanguage%"=="5" GOTO Install2024
IF "%ChoosedLanguage%"=="6" GOTO Install2024Plus
IF "%ChoosedLanguage%"=="7" GOTO Install365
IF "%ChoosedLanguage%"=="10" GOTO RunRemovalTool

ECHO Invalid selection. Please enter a valid option.
TIMEOUT /T 2 >NUL
GOTO SELECT

:Install2019
"Office\2019.exe" /configure "Configuration\2019.xml"
GOTO Done

:Install2019Plus
"Office\2019.exe" /configure "Configuration\2019plus.xml"
GOTO Done

:Install2021
"Office\2021.exe" /configure "Configuration\2021.xml"
GOTO Done

:Install2021Plus
"Office\2021.exe" /configure "Configuration\2021plus.xml"
GOTO Done

:Install2024
"Office\2024.exe" /configure "Configuration\2024.xml"
GOTO Done

:Install2024Plus
"Office\2024.exe" /configure "Configuration\2024plus.xml"
GOTO Done

:Install365
"Office\365.exe"
GOTO Done

:RunRemovalTool
"Office\SetupProd_OffScrub.exe"
GOTO Done

:Done
ECHO.
ECHO Returning to MOPPI main menu...
TIMEOUT /T 3 >NUL
GOTO SELECT

:ExitApp
ECHO Exiting MOPPI...
TIMEOUT /T 1 >NUL
EXIT
