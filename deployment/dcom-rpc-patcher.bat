@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If %2'==' Goto ayuda
If %3'==' Goto ayuda
If Not %4'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ------------------------------------
echo.
echo LAN massive patch DCOM RPC MS03-039
echo.
echo Jordi Corrales
echo Systems Department
echo Example Corp International
echo ------------------------------------
echo.
echo Usage: 
echo.
echo %~n0 IPstart-IPend languaje s.o. reboot
echo.
echo Languajes:               S.O.:
echo.
echo 1. English               1. Windows NT
echo 2. Spanish               2. Windows 2000
echo 3. German                3. Windows XP
echo 4. Portuguese (br)       4. ALL
echo 5. Portuguese (pr)
echo 6. French 
echo 7. Italian 
echo 8. Turkish 
echo.
echo Reboot: Toggle on/off reboot after successful installation
echo.
echo 0. off 
echo 1. on 
echo.
goto end
:Empieza
set rbt=
If %4'==0' set rbt=/z
If %2'==1' set idioma=ENU
If %2'==2' set idioma=ESN
If %2'==3' set idioma=DEU
If %2'==4' set idioma=PTB
If %2'==5' set idioma=PTG
If %2'==6' set idioma=FRA
If %2'==7' set idioma=ITA
If %2'==8' set idioma=TRK
echo.
xfrpcss %1 > scan.txt
goto %3
:1
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c WindowsNT4Workstation-KB824146-x86-%idioma%.exe /q %rbt%) 
goto end

:2
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c Windows2000-KB824146-x86-%idioma%.exe /q %rbt%) 
goto end

:3
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c WindowsXP-KB824146-x86-%idioma%.exe /q %rbt%) 

:4
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c WindowsNT4Workstation-KB824146-x86-%idioma%.exe /q %rbt%) 
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c Windows2000-KB824146-x86-%idioma%.exe /q %rbt%) 
FOR /F "TOKENS=1 DELIMS=[" %%F IN ('TYPE SCAN.TXT ^| FIND /I "VULN"') DO (psexec \\%%F -i -s -c WindowsXP-KB824146-x86-%idioma%.exe /q %rbt%) 
goto end

:END