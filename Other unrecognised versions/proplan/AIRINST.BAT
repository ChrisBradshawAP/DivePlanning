@ECHO OFF
ECHO This software take up approximately 200K bytes on the hard drive.
ECHO Please ensure this space is available before installation.
ECHO If required, hit CTRL-C to cancel the installation and start again
PAUSE
ECHO PRO-DIVE planner is normally installed in directory C:\PRODIVE
ECHO To install PRO-DIVE planner into a different directory, add the
ECHO name of the directory after INSTALL.  To install from drive B:, the 
ECHO directory name on the C: drive must always be specified, with B: after 
ECHO it. eg: 
ECHO  INSTALL  installs from the A: drive into directory C:\PRODIVE
ECHO  INSTALL DIVEPLAN  installs from the A: drive into directory C:\DIVEPLAN
ECHO  INSTALL PRODIVE B:  installs from the B: drive into directory C:\PRODIVE
ECHO If required, hit CTRL-C to cancel the installation and start again
PAUSE
IF "%1"=="" GOTO PRODIR
ECHO Making directory C:\%1, please wait . . .
MD C:\%1
CD C:\%1
GOTO COPSTART
:PRODIR
ECHO Making directory C:\PRODIVE, please wait . . .
MD C:\PRODIVE
cd c:\PRODIVE
GOTO COPSTART
:COPSTART
IF "%2"=="B" GOTO BSTART
IF "%2"=="b" GOTO BSTART
IF "%2"=="B:" GOTO BSTART
IF "%2"=="b:" GOTO BSTART
:START
IF EXIST A:DIVEPLAN.EXE GOTO ACOPY
ECHO Insert your Air/Nitrox PRO-DIVE PLANNER disc in drive A:,
PAUSE
GOTO START
:ACOPY
ECHO Copying files, Please wait . . .
A:INSCOUNT
ECHO.
GOTO END
:BSTART
IF EXIST B:DIVEPLAN.EXE GOTO BCOPY
ECHO Insert your Air/Nitrox PRO-DIVE PLANNER disc in drive B:,
PAUSE
GOTO BSTART
:BCOPY
ECHO Copying files, Please wait . . .
B:INSCOUNT b
ECHO.
:END
C:
