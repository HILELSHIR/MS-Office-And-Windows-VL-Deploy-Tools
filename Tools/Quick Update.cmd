@cls
@echo off

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
) || cd /d "%~dp0"

:::: Run as Admin with native shell, any path, params, loop guard, minimal i/o, by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\" &call \"%%2\" %%3" & set "_= %*"
>nul fltmc || if "%f0%" neq "%~f0" ( cd.>"%tmp%\runas.Admin" & start "%~n0" /high "%tmp%\runas.Admin" "%~f0" "%_:"=""%" &exit /b )

cd..
del /q "D:\Software\MS Tools Pack\All In One\*"
call %windir%\Compress.cmd * "D:\Software\MS Tools Pack\All In One\MS Office And Windows VL Deploy Tools"

cd Tools
set "SetupPath=%cd%"
for /f %%f in ('dir /b d:\setup\*office*C2R*VL*.iso') do call "%SetupPath%\Update ISO Tool" "d:\setup\%%f"

explorer "d:\setup\"
explorer "D:\Software\MS Tools Pack\All In One\"