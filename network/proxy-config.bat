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
echo LAN massive proxy modifier
echo.
echo Jordi Corrales
echo IT Dept - ExampleCorp International
echo ------------------------------------
echo.
echo Usage: 
echo.
echo %~n0 IPstart IPend IPproxy:port
echo.
echo Example: 
echo %~n0 192.168.103.1 192.168.103.255 192.168.103.7:8080
echo.
goto END2

:Empieza
echo Computers OK > results.txt
echo --------------- >> results.txt
echo Pinging ...
FOR /F "TOKENS=1,2,3,4 DELIMS=." %%F IN ('echo %1') DO (set /a a=%%F & set /a b=%%G & set /a c=%%H & set /a d=%%I) 

:d
set /a d=%d% + 1
If "%d%" EQU "256" goto c
echo %a%.%b%.%c%.%d% ...
ping -n 1 -w 500 %a%.%b%.%c%.%d%  > nul
set rsp=%errorlevel%
If %rsp%'==0' Goto cambia

:sigue
If "%a%.%b%.%c%.%d%" EQU "%2" goto end
goto d

:cambia
set ip=%a%.%b%.%c%.%d%
psloggedon.exe \\%ip% -l | find /i "/" > tmp.cwt
FOR /F "TOKENS=3 DELIMS= " %%F IN ('type tmp.cwt') DO (set loged=%%F)
psgetsid \\%ip% %loged% | find /i "S-" > tmp.cwt
FOR /F %%F IN ('type tmp.cwt') DO (set sidusr=%%F)
reg add "\\%ip%\HKU\%sidusr%\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyServer /d %3 /f > nul
set rsp=%errorlevel%
If %rsp%'==0' echo %ip% >> results.txt
del tmp.cwt
goto sigue

:c
set /a c=%c% + 1
set d=0
If "%c%" EQU "255" goto b
goto d

:b
set /a b=%b% + 1
set c=0
If "%b%" EQU "255" goto a
goto d

:a
set /a a=%a% + 1
set a=0
If "%a%" EQU "255" goto end
goto d
goto end

:END
echo.
type results.txt
set a=
set b=
set c=
set d=
set rsp=
set ip=
set sidusr=
set loged=

:END2