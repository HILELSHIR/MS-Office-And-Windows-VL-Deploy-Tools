@cls
@echo off
>nul chcp 437
SETLOCAL EnableDelayedExpansion
title Visual UI 365 Edition

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

set "root="
if exist "%ProgramFiles%\Microsoft Office\root"         set "root=%ProgramFiles%\Microsoft Office\root"
if exist "%ProgramFiles(x86)%\Microsoft Office\root"     set "root=%ProgramFiles(x86)%\Microsoft Office\root"
if not defined root (
    echo.
    echo Error ### Fail to find integrator.exe Tool
    echo.
    pause
    exit /b
)

echo.
echo -- Integrate Mondo 2016 Volume License
"!root!\Integration\integrator" /I /License PRIDName=MondoVolume.16 PidKey=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2

echo -- Clean Registry Keys
for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (
	1>nul 2>&1 reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Microsoft\Office" /f
	1>nul 2>&1 reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Wow6432Node\Microsoft\Office" /f
	1>nul 2>&1 reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides" /f
)

echo -- Install Visual UI Registry Keys
for %%# in (Word, Excel, Powerpoint, Access, Outlook, Publisher, OneNote, project, visio) do (
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" /f /v "%%#" /t REG_SZ /d "{FBDB3E18-A8EF-4FB3-9183-DFFD60BD0984},{CE5FFCAF-75DA-4362-A9CB-00D2689918AA},"
)

echo -- Done.
echo.

echo Note:
echo To initiate the Visual Refresh,
echo it may be required to start some Office apps
echo a couple of times.
echo.
echo Many Thanks to Xtreme21 ^& Krakatoa
echo for helping make and debug this script
echo.

pause
exit /b