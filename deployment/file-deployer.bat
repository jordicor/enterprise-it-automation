@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If %2'==' Goto ayuda
If Not %3'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ------------------------------------
echo.
echo LAN massive file deployer
echo.
echo Jordi Corrales
echo Systems Department
echo Example Corp International
echo ------------------------------------
echo.
echo Usage: 
echo.
echo %~n0 IPstart IPend
echo.
echo Example: 
echo %~n0 172.16.0.1 172.16.15.255
echo.
goto :EOF

:Empieza
set vuelve=Empeza2
set calc=%2
goto conv

:Empeza2
set /a longip2=(256*256*256*%a%)+(256*256*%b%)+(256*%c%)+%d%
set vuelve=Empeza3
set calc=%1
goto conv

:Empeza3
set /a longip=(256*256*256*%a%)+(256*256*%b%)+(256*%c%)+%d%

If "%longip%" GEQ "%longip2%" goto END
echo  Results > results.txt
echo ---------- >> results.txt
echo Pinging ...

:d
If %d% EQU 255 goto c
echo %a%.%b%.%c%.%d% ...
ping -n 1 -w 500 %a%.%b%.%c%.%d%  > nul
set rsp=%errorlevel%
set ip=%a%.%b%.%c%.%d%
set vuelve=d2
set calc=%ip%
goto conv

:d2
set /a longip=(256*256*256*%a%)+(256*256*%b%)+(256*%c%)+%d%
If %rsp% EQU 0 Goto cambia
If %rsp% GEQ 1 echo %ip% [Down] >> results.txt

:sigue
If "%longip%" GEQ "%longip2%" goto end
set /a d+=1
goto d

:cambia
set ip=%a%.%b%.%c%.%d%
set dest=
FOR /F %%F IN ('psexec \\%ip% cmd /c set ^| find /i "windir"') DO (set wind=%%F)
set dest=%wind:~10,10%

IF DEFINED dest (goto copia) else (echo %ip% [Ipc$] >> results.txt)
goto sigue

:copia 
copy /y ip.vbs "\\%ip%\c$\%dest%\system32" > nul
reg add \\%ip%\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v ShowIp /d c:\%dest%\system32\ip.vbs /f > nul
psexec -d \\%ip% cmd /c ip.vbs
set rsp=%errorlevel%
If %rsp% EQU 0 echo %ip% [Done] >> results.txt
If %rsp% GEQ 1 goto fallo
goto sigue

:fallo
echo %ip% [Fail] >> results.txt
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
set ip=
set rsp=
set dest=
set wind=
set dest2=
set vuelve=
goto :EOF

:conv
FOR /F "TOKENS=1,2,3,4 DELIMS=." %%F IN ('echo %calc%') DO (set /a a=%%F & set /a b=%%G & set /a c=%%H & set /a d=%%I) 
goto %vuelve%
