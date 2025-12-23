@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If %2'==' Goto ayuda
If %3'==' Goto ayuda
If %4'==' Goto ayuda
If Not %5'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ________________________
echo.
echo Patcher masivo para LAN
echo por Jordi Corrales
echo 	_________________________
echo.
echo Uso:
echo.
echo %~n0 ipinicio ipfinal patch.exe 0/1
echo.
echo Reiniciar despues de 
echo instalacion satisfactoria: 0 - off , 1 - on
echo. 
echo.
echo Ejemplo:
echo %~n0 172.16.0.1 172.16.15.254 KB824146.exe 0
echo.
goto end
:Empieza
echo Enviando pings..
FOR /F "TOKENS=1,2,3,4 DELIMS=." %%F IN ('echo %1') DO (set /a a=%%F & set /a b=%%G & set /a c=%%H & set /a d=%%I) 
set rbt=
If %4'==0' set rbt=/z
:d
set /a d+=1
If "%d%" EQU "255" goto c
echo %a%.%b%.%c%.%d% ...
ping -n 1 -w 500 %a%.%b%.%c%.%d%  > nul
set rsp=%errorlevel%
If %rsp%'==0' Goto cambia

:sigue
If "%a%.%b%.%c%.%d%" EQU "%2" goto end
goto d

:cambia
set ip=%a%.%b%.%c%.%d%
psexec \\%ip% -i -s -c %3 /q %rbt%
set rsp=%errorlevel%
If %rsp% EQU 3010 echo %ip% %3 [hecho] %rsp% >> resultado.txt
if %rsp% EQU 1603 echo %ip% %3 [fallo] %rsp% >> resultado.txt
If %rsp% EQU 0 echo %ip% %3 [hecho] %rsp% >> resultado.txt
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

:end
echo.
set a=
set b=
set c=
set d=
set rbt=