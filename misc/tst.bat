@echo off

FOR /F "TOKENS=1,2,3,4 DELIMS=." %%F IN ('echo %1') DO (set /a a=%%F & set /a b=%%G & set /a c=%%H & set /a d=%%I) 

set /a longip=(2.56*2.56*2.56*%a%)+(2.56*2.56*%b%)+(2.56*%c%)+%d%

echo.
echo %a% %b% %c% %d%
echo Longip %longip%
echo.

goto end2






set vuelve=Empeza3
set calc=%1
goto conv

:Empeza3
set /a longip=(256*256*256*%a%)+(256*256*%b%)+(256*%c%)+%d%


echo.
echo Longip -> %longip%
echo Longip2 -> %longip2%
echo.
goto :EOF

If "%longip%" GEQ "%longip2%" goto end2
echo  Results > results.txt
echo ---------- >> results.txt
echo Pinging ...
set rbt=/z
If %4'==1' set rbt=


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
psexec \\%ip% -i -s -c %3 /q %rbt%
set rsp=%errorlevel%
If %rsp% EQU 3010 echo %ip% %3 [done] %rsp% >> resultado.txt
if %rsp% EQU 1603 echo %ip% %3 [fail] %rsp% >> resultado.txt
If %rsp% EQU 0 echo %ip% %3 [done] %rsp% >> resultado.txt
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

:end
echo.
echo.
type results.txt
echo.
goto end2
:end2
set a=
set b=
set c=
set d=
set ip=
set rsp=
set rbt=
set vuelve=
goto :EOF


