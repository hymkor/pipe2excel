setlocal
call :"%1"
endlocal
exit /b

:""
    set GOARCH=386
    go build
    exit /b

:"package"
    zip -9 pipe2excel-%DATE:/=%.zip pipe2excel.exe readme.md
    exit /b
