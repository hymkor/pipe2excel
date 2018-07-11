setlocal
call :"%1"
endlocal
exit /b

:""
    set GOARCH=386
    go build
    exit /b

:"package"
    for /f %%I in ('pipe2excel -v') do set VERSION=%%I
    zip -9 pipe2excel-%VERSION%.zip pipe2excel.exe readme.md
    exit /b

:"upgrade"
    @echo off
    for /F %%I in ('where pipe2excel') do if not "%%I" == "%~dp0pipe2excel.exe" copy /-Y /v "%~dp0pipe2excel.exe" %%I
    exit /b
