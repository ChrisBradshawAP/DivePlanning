@ECHO OFF
IF "%1"=="B" GOTO BSTART
IF "%1"=="b" GOTO BSTART
IF "%1"=="B:" GOTO BSTART
IF "%1"=="b:" GOTO BSTART
ECHO AIR/NITROX files to be copied to drive A:
ECHO press 'CTRL-C' to cancel
ECHO Insert a blank disc in drive A:,
PAUSE
IF EXIST A:*.* GOTO :NOTBLANK
ECHO Copying files, Please wait . . .
COPY airnplan.exe A:\diveplan.exe   > NUL
COPY divelog.exe A:\divelog.exe   > NUL
COPY helvb.fon A:\helvb.fon   > NUL
COPY diveplan.cob A:\diveplan.cop   > NUL
COPY msherc.com A:\msherc.com   > NUL
COPY image2.bit A:\image2.bit > NUL
COPY manual.txt A:manual.txt > NUL
COPY plan.bat A:\plan.bat   > NUL
COPY sinstall.exe A:\install.exe > NUL
reccount a
ECHO.
ECHO Air/Nitrox PRO-DIVE PLANNER copied to drive A:
ECHO.
GOTO :END
:NOTBLANK  
ECHO A: disc not blank
GOTO :END
:BSTART
ECHO AIR/NITROX files to be copied to drive B:
ECHO press 'CTRL-C' to cancel
ECHO Insert a blank disc in drive B:,
PAUSE
IF EXIST B:*.* GOTO :NOTBBLANK
ECHO Copying files, Please wait . . .
COPY airnplan.exe B:\diveplan.exe   > NUL
COPY divelog.exe B:\divelog.exe   > NUL
COPY helvb.fon B:\helvb.fon   > NUL
COPY diveplan.cob B:\diveplan.cop   > NUL
COPY msherc.com B:\msherc.com   > NUL
COPY image2.bit B:\image2.bit > NUL
COPY manual.txt B:manual.txt > NUL
COPY plan.bat B:\plan.bat   > NUL
COPY sinstall.exe B:\install.exe > NUL
reccount a
ECHO.
ECHO Air/Nitrox PRO-DIVE PLANNER copied to drive B:
ECHO.
GOTO :END
:NOTBBLANK  
ECHO B: disc not blank
:END
