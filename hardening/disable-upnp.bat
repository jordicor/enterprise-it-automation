@echo off
If %1'==/?' Goto ayuda
If %1'==' Goto ayuda
If Not %2'==' Goto ayuda
goto Empieza
:ayuda
echo.
echo ------------------------------------
echo Disable UPNP
echo by CrowDat Kurobudetsu
echo ------------------------------------
echo.
echo Usage: 
echo.
echo %~n0 file.txt
echo.
echo Example: 
echo %~n0 results.txt
echo.
goto end

:Empieza
FOR /F "TOKENS=1 DELIMS= " %%F IN ('type %1 ^| find /i "." ') DO (
echo Desactivando upnp en %%F ...

reg add \\%%F\HKLM\SYSTEM\CurrentControlSet\Services\upnphost /v start /t REG_DWORD /d 4 /f
reg add \\%%F\HKLM\SYSTEM\CurrentControlSet\Services\SSDPSRV /v start /t REG_DWORD /d 4 /f > nul
)


:END

