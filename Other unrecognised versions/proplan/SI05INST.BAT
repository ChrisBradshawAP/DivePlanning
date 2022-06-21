@echo off
rem ** - SI05INST.BAT - THIS BATCH FILE IS CALLED FROM INSTALL.EXE 
rem **
rem ** -  %1 = hard drive letter (c):
rem ** -  %2 = target directory on hard drive (si05)
rem ** -  %3 = source drive letter (a):
rem ** 
rem ** - IT MAY BE INVOKED DIRECT FROM DOS AS LONG AS THE ABOVE PARAMETERS 
rem ** - ARE USED CORRECTLY.  
rem **
:TOP
if "%1" == "" goto ER1 
if "%2" == "" goto ER1
if "%3" == "" goto ER1

goto START

:ER1
cls
echo.
echo.
echo.
echo.
echo.
echo ERROR INVALID PARAMETERS - type "INSTALL" to install
echo.
echo.
echo.
goto END

:START
%1:
cd \ > nul
if not exist %1:\%2\nul md %2 > nul
cd %2 > nul
copy %3:. > nul
pkunzip -o si05 > nul

:END
