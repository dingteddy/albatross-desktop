@echo off
@echo current path: %~dp0
::reg add HKEY_CURRENT_USER\Software\teamtop3\WindSeal /v aaa /t REG_SZ /d %~dp0\abt.exe /f
reg add HKEY_CURRENT_USER\Software\TortoiseSVN\DiffTools /v .xlsx /t REG_SZ /d "\"C:\Program Files (x86)\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE\" %%base %%mine \"%~dp0\config.ini\"" /f
reg add HKEY_CURRENT_USER\Software\TortoiseSVN\MergeTools /v .xlsx /t REG_SZ /d "%~dp0\albatross-desktop.exe %%merged %%theirs %%mine %%base \"%~dp0\config.ini\"" /f
@echo finished!
pause