cd %~dp0
copy /Y ..\game-of-life.xlsm .\bin\
cscript vbac.wsf decombine
move /Y .\src\game-of-life.xlsm\Sheet1.dcm ..\code.vb
del .\bin\game-of-life.xlsm

