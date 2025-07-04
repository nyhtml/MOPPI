@ECHO OFF && COLOR 1F && TITLE MOPPI v2025.06

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
ECHO    Copyright (c) 2013-%DATE:~6,4% Stephan Pringle. All rights reserved.
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
SET /p ChoosedLanguage=Enter a number and press ENTER key or 0 to quit: 

IF %ChoosedLanguage% == 0 GOTO E
IF %ChoosedLanguage% GEQ 1 IF %ChoosedLanguage% LEQ 187 GOTO %ChoosedLanguage%

GOTO SELECT

:1
"Office\2019.exe" /configure "Configuration\2019.xml"&GOTO DONE
:2
"Office\2019.exe" /configure "Configuration\2019plus.xml"&GOTO DONE
:3
"Office\2021.exe" /configure "Configuration\2021.xml"&GOTO DONE
:4
"Office\2021.exe" /configure "Configuration\2021plus.xml"&GOTO DONE
:5
"Office\2024.exe" /configure "Configuration\2024.xml"&GOTO DONE
:6
"Office\2024.exe" /configure "Configuration\2024plus.xml"&GOTO DONE
:7
"Office\365.exe"&GOTO DONE
:10
"Office\SetupProd_OffScrub.exe"&GOTO DONE

:DONE
GOTO SELECT
:E
