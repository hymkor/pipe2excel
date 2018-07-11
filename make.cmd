setlocal
set "PROMPT=$G "
set EXE=pipe2excel.exe
for %%I in (%EXE%) do set "NAME=%%~nI"
call :"%1"
endlocal
exit /b

:""
    set GOARCH=386
    go build
    exit /b

:"package"
    for /f %%I in ('%NAME% -v') do set VERSION=%%I
    zip -9 "%NAME%-%VERSION%.zip" "%EXE%" readme.md
    exit /b

:"upgrade"
    for /f %%I in ('where %EXE%') do if not "%%I" == "%~dp0%EXE%" copy /v /-y "%~dp0%EXE%" "%%I"
    exit /b
