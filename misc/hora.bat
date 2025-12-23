@echo off
FOR /F "TOKENS=1 DELIMS=:" %%F IN ('time /t') DO (if %%F GEQ 6 if %%F LEQ 22 echo Hola!)

