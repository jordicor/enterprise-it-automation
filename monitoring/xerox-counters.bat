@echo off
cls
echo GET /properties/billingCounters.dhtml HTTP/1.1>%temp%pide.txt
echo Host: 172.28.4.56>>%temp%pide.txt
echo User-Agent: Mozilla/5.0 (Windows; U; Windows NT 5.1; es-ES; rv:1.7.8) Gecko/20050511 Firefox/1.0.4>>%temp%pide.txt
echo Accept: text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5>>%temp%pide.txt
echo Accept-Language: es-es,es;q=0.5>>%temp%pide.txt
echo Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7>>%temp%pide.txt
echo Referer: http://172.28.4.56/displayTree.dhtml>>%temp%pide.txt
echo Cookie: statusSelected=2; propSelected=6; propHierarchy=1000000000000>>%temp%pide.txt
echo Cache-Control: max-age=0 >>%temp%pide.txt
echo.>>%temp%pide.txt
echo.>>%temp%pide.txt
echo.>>%temp%pide.txt
echo HELO Fico> %temp%smtp.txt
echo MAIL FROM: xeroxreports@example-corp.com>> %temp%smtp.txt
echo RCPT TO: report_xerox_mollet@example-corp.com>> %temp%smtp.txt
echo DATA>> %temp%smtp.txt
echo To: Reportes Xerox Mollet ^< report_xerox_mollet@example-corp.com^>>> %temp%smtp.txt
echo Subject: Reporte Impresiones Realizadas>> %temp%smtp.txt
echo MIME-Version: 1.0>> %temp%smtp.txt
echo X-Mailer: CwT Batch Script 0.69b>> %temp%smtp.txt
echo From: Reporte Impresiones Xerox %2 ^<xeroxreports@example-corp.com^>>> %temp%smtp.txt
echo.>> %temp%smtp.txt
echo Proceso iniciado
echo Recuperando datos en [I58SUW74] ...
call :copias 172.28.4.41 I58SUW74 310-873804 XeroxWorkCentrePro-40C 1
echo ok
echo Recuperando datos en [I58SUW75] ...
call :copias 172.28.4.42 I58SUW75 310-873979 XeroxWorkCentrePro-32C 1
echo ok
echo Recuperando datos en [I58SUW76] ...
call :copias 172.28.4.43 I58SUW76 2232221642 XeroxWorkCentrePro-35 0
echo ok
echo Recuperando datos en [I58SUW77] ...
call :copias 172.28.4.44 I58SUW77 310-873758 XeroxWorkCentrePro-32C 1
echo ok
echo Recuperando datos en [I58SUW78] ...
call :copias 172.28.4.45 I58SUW78 2232216606 XeroxWorkCentrePro-45 0
echo ok
echo Recuperando datos en [I58SUW70] ...
call :copias 172.28.4.47 I58SUW70 2232216789 XeroxWorkCentrePro-45 0
echo ok
echo Recuperando datos en [I58SUW71] ...
call :copias 172.28.4.48 I58SUW71 2232155704 XeroxWorkCentrePro-55 0
echo ok
echo Recuperando datos en [I58SUW72] ...
call :copias 172.28.4.49 I58SUW72 2232187622 XeroxWorkCentrePro-55 0
echo ok
echo Recuperando datos en [I58SUW86] ...
call :copias 172.28.4.53 I58SUW86 2232221430 XeroxWorkCentrePro-35 0
echo ok
echo Recuperando datos en [I58SUW88] ...
call :copias 172.28.4.55 I58SUW88 310-873967 XeroxWorkCentrePro-40C 1
echo ok
echo Recuperando datos en [I58SUW89] ...
call :copias 172.28.4.56 I58SUW89 310-873969 XeroxWorkCentrePro-32C 1
echo ok
echo .>> %temp%smtp.txt
echo.>> %temp%smtp.txt
echo quit>> %temp%smtp.txt
echo Enviando informacion por smtp...
nc 172.28.2.5 25 < %temp%smtp.txt>nul
echo ok
echo Proceso Finalizado
del %temp%pide.txt>nul
del %temp%smtp.txt>nul
del %temp%suelta.txt>nul
goto :EOF

:copias
nc -w 1 %1 80 < %temp%pide.txt > %temp%suelta.txt
echo Nombre: "%2";>>%temp%smtp.txt
echo Modelo: "%4";>>%temp%smtp.txt
echo N�Serie: "%3";>>%temp%smtp.txt
If %5'==1' Goto wcolor
:cnt
FOR /F "TOKENS=2 DELIMS==" %%F IN ('type %temp%suelta.txt ^| find "var markedBWandColorImages" ') DO (echo Todas las impresiones %%F>> %temp%smtp.txt)
echo ------------->>%temp%smtp.txt
goto :EOF

:wcolor
FOR /F "TOKENS=2 DELIMS==" %%F IN ('type %temp%suelta.txt ^| find "var markedBWImages" ') DO (echo Impresiones en B/N %%F>> %temp%smtp.txt)
FOR /F "TOKENS=2 DELIMS==" %%F IN ('type %temp%suelta.txt ^| find "var markedColorImages" ') DO (echo Impresiones en Color %%F>> %temp%smtp.txt)
goto cnt
