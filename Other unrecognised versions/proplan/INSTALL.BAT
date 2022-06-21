@ECHO OFF
ECHO This installation takes up approximately 500K bytes of disc space.
ECHO Make sure you have enough free disc space before starting
ECHO If required, hit CTRL-C to cancel the installation and start again
PAUSE
IF "%2"=="B" GOTO BSTART
IF "%2"=="b" GOTO BSTART
IF "%2"=="B:" GOTO BSTART
IF "%2"=="b:" GOTO BSTART
:START
IF EXIST A:DIVEPLAN.EXE GOTO ACOPY
ECHO Insert your installation PRO-DIVE PLANNER disc in drive A:,
PAUSE
GOTO START
:ACOPY
ECHO Copying files, Please wait . . .
A:sinstall a
ECHO.
GOTO END
:BSTART
IF EXIST B:DIVEPLAN.EXE GOTO BCOPY
ECHO Insert your installation PRO-DIVE PLANNER disc in drive B:,
PAUSE
GOTO BSTART
:BCOPY
ECHO Copying files, Please wait . . .
B:sinstall b
ECHO.
:END
C:
MD C:\PRODEAL
CD C:\PRODEAL
IF "%2"=="B" GOTO BSTART2
IF "%2"=="b" GOTO BSTART2
IF "%2"=="B:" GOTO BSTART2
IF "%2"=="b:" GOTO BSTART2
ECHO DEALER files to be copied to directory C:\PRODEAL
ECHO Copying files, Please wait . . .
COPY A:airnplan.exe C:airnplan.exe   > NUL
COPY A:trmxplan.exe C:trmxplan.exe   > NUL
COPY A:rebrplan.exe C:rebrplan.exe   > NUL
COPY A:reccount.exe C:reccount.exe > NUL
COPY A:sellcoun.t C:sellcoun.t > NUL
COPY A:sinstall.exe C:sinstall.exe > NUL
COPY A:helvb.fon C:helvb.fon   > NUL
COPY A:diveplan.cob C:diveplan.cob   > NUL
COPY A:makeair.bat C:makeair.bat   > NUL
COPY A:maketri.bat C:maketri.bat   > NUL
COPY A:makereb.bat C:makereb.bat   > NUL
COPY A:msherc.com C:msherc.com   > NUL
COPY A:plan.bat C:plan.bat   > NUL
COPY A:manual.txt C:manual.txt > NUL
COPY A:dealer.bat C:dealer.bat   > NUL
COPY A:image2.bit C:image2.bit > NUL
COPY A:readme.bat C:readme.bat   > NUL
ECHO.
GOTO END2
:BSTART2
ECHO DEALER files to be copied to directory C:\PRODEAL
ECHO Copying files, Please wait . . .
COPY B:airnplan.exe C:airnplan.exe   > NUL
COPY B:trmxplan.exe C:trmxplan.exe   > NUL
COPY B:rebrplan.exe C:rebrplan.exe   > NUL
COPY B:reccount.exe C:reccount.exe > NUL
COPY B:sellcoun.t C:sellcoun.t > NUL
COPY B:sinstall.exe C:sinstall.exe > NUL
COPY B:helvb.fon C:helvb.fon   > NUL
COPY B:diveplan.cob C:diveplan.cob   > NUL
COPY B:makeair.bat C:makeair.bat   > NUL
COPY B:maketri.bat C:maketri.bat   > NUL
COPY B:makereb.bat C:makereb.bat   > NUL
COPY B:msherc.com C:msherc.com   > NUL
COPY B:manual.txt C:manual.txt > NUL
COPY B:plan.bat C:plan.bat   > NUL
COPY B:dealer.bat C:dealer.bat   > NUL
COPY B:image2.bit C:image2.bit > NUL
COPY B:readme.bat C:readme.bat   > NUL
ECHO.
GOTO END2
:END2
ECHO Dealer PRO-DIVE PLANNER installed in directory C:\PRODEAL
ECHO Type DEALER to see menu, or README for more explanation
ECHO.
C:
PAUSE
CALL README.BAT
CALL DEALER
