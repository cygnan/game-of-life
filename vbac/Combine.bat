cd %~dp0
copy /Y ..\code.vb .\src\game-of-life.xlsm\Sheet1.dcm
cscript vbac.wsf combine
move /Y .\bin\game-of-life.xlsm ..\
del .\src\game-of-life.xlsm\Sheet1.dcm
