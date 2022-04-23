@cls
@echo off
>nul chcp 437
setlocal enabledelayedexpansion

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
) || cd /d "%~dp0"

rem Rebuild the Tokens.dat file
rem https://docs.microsoft.com/en-us/windows-server/get-started/activation-rebuild-tokens-dat-file

echo.
echo Gain Administrator rights

:::: Run as Admin with native shell, any path, params, loop guard, minimal i/o, by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\" &call \"%%2\" %%3" & set "_= %*"
>nul fltmc || if "%f0%" neq "%~f0" ( cd.>"%tmp%\runas.Admin" & start "%~n0" /high "%tmp%\runas.Admin" "%~f0" "%_:"=""%" &exit /b )

echo Stop SPP Service
1>nul 2>nul net stop sppsvc

echo Remove Old [Tokens.dat] file
1>nul 2>nul pushd "%windir%\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
)
1>nul 2>nul pushd "%windir%\System32\spp\store\" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
)
1>nul 2>nul pushd "%windir%\System32\spp\store\2.0\" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
)
1>nul 2>nul pushd "%windir%\System32\spp\store_test\2.0\" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
)

echo Start SPP Service
>nul net start sppsvc

echo Rebuild [Tokens.dat] file
>nul cscript.exe %windir%\system32\slmgr.vbs /rilc

timeout 5
exit /b