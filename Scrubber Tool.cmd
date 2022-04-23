@cls
@echo off
>nul chcp 437

set params=
set cscript=
set result="%temp%\result"
SETLOCAL EnableDelayedExpansion
title Scrubber Tool

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
)

:::: Run as Admin with native shell, any path, params, loop guard, minimal i/o, by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\" &call \"%%2\" %%3" & set "_= %*"
>nul fltmc || if "%f0%" neq "%~f0" ( cd.>"%tmp%\runas.Admin" & start "%~n0" /high "%tmp%\runas.Admin" "%~f0" "%_:"=""%" &exit /b )
	
cd /d "%~dp0"

echo:
echo :: Clean Office MSI ::
echo * 2003 (11)
echo * 2007 (12)
echo * 2010 (14)
echo * 2013 (15)
echo * 2016 (16)
echo:
echo :: Clean Office C2R ::
echo * 2016
echo * 2019
echo * 2021

echo.
pause

cls
echo.
call :Scrubb

:end
echo.
echo please wait 5 seconds..
SETLOCAL DisableDelayedExpansion
PING -n 5 127.0.0.1 >NUL
goto :eof

:Scrubb
if exist "%windir%\SysWOW64\cscript.exe" set cscript="%windir%\SysWOW64\cscript.exe"
if exist "%windir%\system32\cscript.exe" set cscript="%windir%\system32\cscript.exe"

echo Process :: Clean Keys ^& Licences
1>nul 2>&1 Tools\x32\cleanospp.exe
!cscript! Tools\vbs\OLicenseCleanup.vbs //nologo //b /QUIET
echo ....................................................................................................
echo Process :: Clean Registry ^& Folders
>!result! 2>&1 dir "%ProgramFiles%\Microsoft Office*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles%\%%#"
)
>!result! 2>&1 dir "%ProgramFiles(x86)%\Microsoft Office*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles(x86)%\%%#"
)
>!result! 2>&1 dir "%programfiles%\Common Files\microsoft shared\OFFICE*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%programfiles%\Common Files\microsoft shared\%%#"
)
>!result! 2>&1 dir "%ProgramFiles(x86)%\Common Files\microsoft shared\OFFICE*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles(x86)%\Common Files\microsoft shared\%%#"
)
>nul 2>&1 del /q !result!
for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (set "GUID=S-%%A-%%B-%%C-%%D-%%E-%%F-%%G")
for %%$ in (HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER,HKEY_USERS) do (
	echo "%%$" |>nul find /i "HKEY_USERS" && (
		2>nul reg query "%%$\!GUID!\SOFTWARE\Microsoft" /f office 			  | >>!result! find /i "%%$"
		2>nul reg query "%%$\!GUID!\SOFTWARE\Wow6432Node\Microsoft" /f office | >>!result! find /i "%%$"
	) || (
		2>nul reg query "%%$\SOFTWARE\microsoft" /f office 					  | >>!result! find /i "%%$"
		2>nul reg query "%%$\SOFTWARE\WOW6432Node\microsoft" /f office 		  | >>!result! find /i "%%$"
	)
)
if exist !result! (
	for /f "tokens=*" %%$ in ('type !result!') do (
		>nul 2>&1 reg delete "%%$" /f
	)
	>nul 2>&1 del /q !result!
)
echo ....................................................................................................
echo Process :: OffScrubC2R.vbs
!cscript! Tools\vbs\OffScrubC2R.vbs //nologo //b ALL /NoCancel /Force /OSE /Quiet /NoReboot /Passive
echo ....................................................................................................
for %%G in (OffScrub_O16msi.vbs,OffScrub_O15msi.vbs,OffScrub10.vbs,OffScrub07.vbs,OffScrub03.vbs) do (
	echo Process :: %%G
	!cscript! Tools\vbs\%%G //nologo //b ALL /NoCancel /Force /OSE /Quiet /NoReboot /Passive
	echo.
)
echo Process :: Clean Leftover
>!result! 2>&1 dir "%ProgramFiles%\Microsoft Office*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles%\%%#"
)
>!result! 2>&1 dir "%ProgramFiles(x86)%\Microsoft Office*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles(x86)%\%%#"
)
>!result! 2>&1 dir "%programfiles%\Common Files\microsoft shared\OFFICE*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%programfiles%\Common Files\microsoft shared\%%#"
)
>!result! 2>&1 dir "%ProgramFiles(x86)%\Common Files\microsoft shared\OFFICE*" /ad /b && (
	for /f "tokens=*" %%# in ('type !result!') do >nul 2>&1 call :DestryFolder "%ProgramFiles(x86)%\Common Files\microsoft shared\%%#"
)
>nul 2>&1 del /q !result!
for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (set "GUID=S-%%A-%%B-%%C-%%D-%%E-%%F-%%G")
for %%$ in (HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER,HKEY_USERS) do (
	echo "%%$" |>nul find /i "HKEY_USERS" && (
		2>nul reg query "%%$\!GUID!\SOFTWARE\Microsoft" /f office 			  | >>!result! find /i "%%$"
		2>nul reg query "%%$\!GUID!\SOFTWARE\Wow6432Node\Microsoft" /f office | >>!result! find /i "%%$"
	) || (
		2>nul reg query "%%$\SOFTWARE\microsoft" /f office 					  | >>!result! find /i "%%$"
		2>nul reg query "%%$\SOFTWARE\WOW6432Node\microsoft" /f office 		  | >>!result! find /i "%%$"
	)
)
if exist !result! (
	for /f "tokens=*" %%$ in ('type !result!') do (
		>nul 2>&1 reg delete "%%$" /f
	)
	>nul 2>&1 del /q !result!
)

goto :eof

:DestryFolder
>nul 2>&1 rd/s/q "%temp%"
>nul 2>&1 md "%temp%"
set "targetFolder=%*"
if exist %targetFolder% (
	rd /s /q %targetFolder%
	if exist %targetFolder% (
		for /f "tokens=*" %%g in ('dir /b/s /a-d %targetFolder%') do move /y "%%g" "%temp%"
		rd /s /q %targetFolder%
	)
)
goto :eof