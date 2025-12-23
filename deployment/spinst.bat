@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If %2'==' Goto ayuda
If %3'==' Goto ayuda
If %4'==' Goto ayuda
If %5'==' Goto ayuda
If %6'==' Goto ayuda
If %7'==' Goto ayuda
If Not %8'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ------------------------------------
echo.
echo LAN massive service pack installer
echo.
echo Jordi Corrales
echo Systems Department
echo Example Corp International
echo ------------------------------------
echo.
echo Usage: 
echo.
echo "%~n0 IPstart IPend 0/1 Domain User Password [\\path\]file.ext"
echo.
echo Toggle on/off reboot after successful installation : 0 - off , 1 - on
echo. 
echo.
echo Example: 
echo %~n0 172.16.0.1 172.16.15.255 0 Domain Administrator 6969 \\SRVAPPS\WIN2KSP4\W2KSP4_EN.EXE
echo.
goto end
:Empieza
echo %1 > tmp.cwt
echo Pinging ...
FOR /F "TOKENS=1 DELIMS=." %%F IN ('type tmp.cwt') DO (set a=%%F) 
FOR /F "TOKENS=2 DELIMS=." %%F IN ('type tmp.cwt') DO (set b=%%F) 
FOR /F "TOKENS=3 DELIMS=." %%F IN ('type tmp.cwt') DO (set c=%%F) 
FOR /F "TOKENS=4 DELIMS=." %%F IN ('type tmp.cwt') DO (set d=%%F) 
set rbt=
If %3'==0' set rbt=/z
:d
set /a d=%d% + 1
If "%d%" EQU "256" goto c
echo %a%.%b%.%c%.%d% ...
ping -n 1 -w 500 %a%.%b%.%c%.%d% | find /i "TTL" >> tmp2.cwt
If "%a%.%b%.%c%.%d%" EQU "%2" goto go
goto d

:c
set /a c=%c% + 1
set d=0
If "%c%" EQU "256" goto b
goto d

:b
set /a b=%b% + 1
set c=0
If "%b%" EQU "256" goto a
goto d

:a
set /a a=%a% + 1
set a=0
If "%a%" EQU "256" goto end
goto d
goto end
:go
echo Patching ...
FOR /F "TOKENS=3 DELIMS= " %%F IN ('type tmp2.cwt') DO (echo %%F >> tmp3.cwt)
FOR /F "TOKENS=1 DELIMS=:" %%F IN ('type tmp3.cwt') DO (psexec \\%%F -i -s -c runasp.exe /domain:%4 /user:%5 /password:%6 /silent /command:"%7 /q %rbt%") 
del tmp?.cwt > nul
:END