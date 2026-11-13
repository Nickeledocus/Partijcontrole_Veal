@echo off
echo Partijcontrole Veal wordt gestart. Foutmeldingen verschijnen in dit venster.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Partijcontrole-Veal.ps1"
echo.
echo Het programma is afgesloten. Staat hierboven een rode foutmelding? Kopieer die tekst.
pause

