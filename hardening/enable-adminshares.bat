@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If %2'==' Goto ayuda
If Not %3'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ------------------------------------
echo Enable ipc$ admin$ c$
echo by CrowDat Kurobudetsu
echo ------------------------------------
echo.
echo Usage: 
echo.
echo %~n0 IPstart IPend
echo.
echo Example: 
echo %~n0 192.168.103.1 192.168.103.255
echo.
goto END2

:Empieza
FOR /F "TOKENS=1,2,3,4 DELIMS=." %%F IN ('echo %1') DO (set /a a=%%F & set /a b=%%G & set /a c=%%H & set /a d=%%I) 
If "%a%.%b%.%c%.%d%" GEQ "%2" goto END2
echo  Results > results.txt
echo ---------- >> results.txt
echo Pinging ...

:d
If %d% EQU 255 goto c
echo %a%.%b%.%c%.%d% ...
ping -n 1 -w 500 %a%.%b%.%c%.%d%  > nul
set rsp=%errorlevel%
set ip=%a%.%b%.%c%.%d%
If %rsp% EQU 0 Goto cambia
If %rsp% GEQ 1 echo %ip%  [Down] >> results.txt

:sigue
If "%a%.%b%.%c%.%d%" GEQ "%2" goto end
set /a d+=1
goto d

:cambia
set ip=%a%.%b%.%c%.%d%
reg add \\%ip%\HKLM\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters /v autosharewks /t REG_DWORD /d 1 /f > nul
set rsp=%errorlevel%
If %rsp% EQU 0 echo %ip%  [Done] >> results.txt
If %rsp% GEQ 1 echo %ip%  [Fail] >> results.txt
goto sigue

:c
set /a c+=1
set d=0
If "%c%" EQU "255" goto b
goto d

:b
set /a b+=1
set c=0
If "%b%" EQU "255" goto a
goto d

:a
set /a a+=1
set a=0
If "%a%" EQU "255" goto end
goto d
goto end

:END
echo.
echo.
type results.txt
echo.
set a=
set b=
set c=
set d=
set rsp=
set ip=

:END2