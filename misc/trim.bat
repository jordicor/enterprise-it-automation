@echo off
:trimtest
  set test="  ab c    def 123    "
  set test=%test:"=%
  echo [%test%]
  call :trimvar test
  echo [%test%]
  goto :EOF
::::::::::::::::

:: Trim leading and trailing spaces in an environment variable.
:: 1st (and only) parameter is the name of an evironment variable
:: that holds the unquoted value to be trimmed in place

:trimvar
  call :ltrimvar %1
  call :rtrimvar %1
  goto :EOF

:: Trim leading spaces in an environment variable.
:: 1st (and only) parameter is the name of an evironment variable
:: that holds the unquoted value to be trimmed in place

:ltrimvar
  for /f "tokens=1* delims==" %%y in ('set %1') do (
    set ltv_trimbuf=%%z)
  :chkltdone
    if     "%ltv_trimbuf%"==" "      goto ltdone
    if not "%ltv_trimbuf:~0,1%"==" " goto ltdone
    set ltv_trimbuf=%ltv_trimbuf:~1%
    goto chkltdone
  :ltdone
  set %1=%ltv_trimbuf%
  set ltv_trimbuf=
  goto :EOF

:: Trim trailing spaces in an environment variable.
:: 1st (and only) parameter is the name of an evironment variable
:: that holds the unquoted value to be trimmed in place

:rtrimvar
  for /f "tokens=1* delims==" %%y in ('set %1') do (
    set rtv_trimbuf=%%z)
  call :varlength rtv_end rtv_trimbuf
  :chkrtdone
    set /a rtv_end-=1
    if     "%rtv_trimbuf%"==" "      goto rtdone
    call :exec set rtv_lastch=%%rtv_trimbuf:~%rtv_end%%%
    if not "%rtv_lastch%"==" " goto rtdone
    call :exec set rtv_trimbuf=%%rtv_trimbuf:~0,%rtv_end%%%
    goto chkrtdone
  :rtdone
  set %1=%rtv_trimbuf%
  set rtv_trimbuf=
  set rtv_end=
  set rtv_lastch=
  goto :EOF

:: get length of environment variable (via brute force)
:: %1 name of var to hold found length
:: %2 name of var to check (stored value can't contain dbl quotes)

:varlength
  for /f "tokens=1* delims==" %%y in ('set %2') do (
    set vl_val=%%z)
  set vl_vlen=0
  :chklendone
    if "%vl_val%"=="" goto lendone
    set vl_val=%vl_val:~1%
    set /a vl_vlen+=1
    goto chklendone
  :lendone
  set %1=%vl_vlen%
  set vl_vlen=
  set vl_val=
  goto :EOF

:exec
  %*
  goto :EOF
