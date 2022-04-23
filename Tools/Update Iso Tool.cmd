@cls
@echo off
>nul chcp 437

Title Update Iso Tool
setlocal enabledelayedexpansion

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
set sFile=
set sFile=%~f1
set tFolder="%cd%"

echo.

echo --- Check arguments
if not defined sFile (
	echo.
	echo ERROR ### Missing Args
	echo.&pause
	exit /b
)

echo --- Clean Office folder
if exist Office (
	1>nul 2>&1 rd/s/q office
	if exist Office (
		echo ERROR ### fail to clean office folder
		echo.&pause
		exit /b
	)
)

echo --- Set ISO file :: !sFile!
set sFile="%~f1"
if not exist !sFile! (
	echo.
	echo ERROR ### File not exist
	echo.&pause
	exit /b
)
echo !sFile! | 1>nul 2>&1 find /i ".iso" || (
	echo.
	echo ERROR ### Not an iso file
	echo.&pause
	exit /b
)

if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set "7z=tools\x32\7z")
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set "7z=tools\x64\7z")
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set "7z=tools\x64\7z"
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set "7z=tools\x64\7z"

echo --- Extract Office folder From ISO
1>nul 2>&1 "!7z!" x -y -o%tFolder% %sFile% "Office" || (
	rd /s/q office
	echo.
	echo ERROR ### 7Z Extract Failed
	1>nul 2>&1 rd/s/q Office
	echo.&pause
	exit /b
)

if not exist Office (
	echo.
	echo ERROR ### 7Z Extract Failed
	echo.&pause
	exit /b
)

echo --- Verify Office Folder HASH
>"%temp%\tmp" call "tools\Verify Digital Signatures.cmd" Office
type "%temp%\tmp" |>nul find /i "OK!" || (
	echo.
	echo ERROR ### Hash Check Failed
	1>nul 2>&1 rd/s/q Office
	echo.&pause
	exit /b
)

if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set sBit=32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set sBit=64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set sBit=64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set sBit=64

set iso_label=%~n1
echo %iso_label% | >nul find /i "x32-x64" && (set iso_label=%iso_label:~0,29%) || (set iso_label=%iso_label:~0,25%)

echo --- Delete ISO :: !sFile!
1>nul 2>&1 del /q %~f1
if exist %~f1 (
	echo.
	echo ERROR ### Fail to delete :: %~f1
	1>nul 2>&1 rd/s/q Office
	echo.&pause
	exit /b
)

echo --- Create ISO :: !sFile!
1>nul 2>&1 Tools\x!sBit!\oscdimg.exe -m -u1 -L!iso_label! "%cd%" %~f1 || (
	echo.
	echo ERROR ### ISO Creation failed
	1>nul 2>&1 rd/s/q Office
	echo.&pause
	exit /b
)

1>nul 2>&1 rd/s/q Office
echo --- Done ^& Done
echo.
echo Have a Great Day ...
1>nul 2>&1 rd/s/q office
timeout 5
exit /b