echo off
c:
copy a:explorer.bat c:
md si05
cd\si05
copy a:. > nul
pkunzip -o si05 > nul
cd\
cls
@echo  "Type Explorer at the Root Directory to start Programme"


