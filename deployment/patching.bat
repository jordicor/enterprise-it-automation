@echo off
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c Windows2000-KB823980-x86-ENG.exe /q /z)
