@cls
@echo off
>nul chcp 437

set "wmicTool="
title Office Info Tool
setLocal EnableExtensions EnableDelayedExpansion

if /i '%3' EQU '' goto :USAGE
if /i '%2' EQU '' goto :USAGE
if /i '%1' EQU '' goto :USAGE

set "pFilter=%1"
set "pName=%2"
set "pSerial=%3"
set "pClean=%4"

echo.
if exist "%ProgramFiles%\Microsoft Office\root" set "root=%ProgramFiles%\Microsoft Office\root"
if exist "%ProgramFiles%\Microsoft Office\root" set "OSPP=cscript //B //Nologo "%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs""
if exist "%ProgramFiles%\Microsoft Office\root" set "OSPP2=cscript //Nologo "%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs""

if exist "%ProgramFiles(x86)%\Microsoft Office\root" set "root=%ProgramFiles(x86)%\Microsoft Office\root"
if exist "%ProgramFiles(x86)%\Microsoft Office\root" set "OSPP=cscript //B //Nologo "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs""
if exist "%ProgramFiles(x86)%\Microsoft Office\root" set "OSPP2=cscript //Nologo "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs""

if not defined root echo ### ERROR Missing files & echo: & goto :END
if not exist "!Root:~0,-5!\Office16\OSPP.vbs" echo ### ERROR Missing files & echo: & goto :END

set "key="
set pKey="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
>nul 2>nul reg query !pKey! /v ProductReleaseIds && set "key=!pKey!"
set pKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
>nul 2>nul reg query !pKey! /v ProductReleaseIds && set "key=!pKey!"
if not defined key echo ### ERROR Missing registry key & echo: & goto :END

set "sKey="
set "p16x="
set "pType="
set "PidKey="
set "searchString="
call :Validate "!pSerial!"
(echo '%pName%' | >nul find /i "2021") || (echo '%pName%' | >nul find /i "2019") || set "p16x=*"
echo '%pName%' | >nul find /i "volume" && set "pType=Volume" && set "searchString=*%pName:~0,-6%*VL*.xrm-ms"
echo '%pName%' | >nul find /i "retail" && set "pType=Retail" && set "searchString=*%pName:~0,-6%*Retail*.xrm-ms"
echo '%pName%' | >nul find /i "VL"     && set "pType=Volume" && set "searchString=*%pName%*.xrm-ms"
if not defined pType if /i "%pName:~-1%" EQU "R" if /i "$%pName:~4,1%" NEQ "$" (
                                          set "pType=Retail" && set "searchString=*%pName%*.xrm-ms")
										  
if defined pType (

	REM ****************
	REM  # MANUAL WAY #
	REM ****************
	
	if defined pClean if /i '%pClean%' EQU '-Clean' call :Clean_A
	
	for /f "tokens=*" %%$ in ('"2>nul dir /b "!root!\Licenses16\client*.xrm-ms""') do (
		echo Install License :: "!root!\Licenses16\%%$"
		%ospp% /inslic:"!root!\Licenses16\%%$"
	)
	
	echo:
	for /f "tokens=*" %%$ in ('"2>nul dir /b "!root!\Licenses16\!searchString!""') do (
		if defined p16x (
			(echo '%%$' | >nul find /i "2021") || (echo '%%$' | >nul find /i "2019") || (
				(echo '%%$' | >nul find /i "Preview") || (
					echo Install License :: "!root!\Licenses16\%%$"
					%ospp% /inslic:"!root!\Licenses16\%%$"
				)
			)
		) else (
			(echo '%%$' | >nul find /i "Preview") || (
				echo Install License :: "!root!\Licenses16\%%$"
				%ospp% /inslic:"!root!\Licenses16\%%$"
			)
		)
	)
	
	echo:
	if defined pClean if /i '%pClean%' EQU '-Clean' call :Clean_M
	if defined sKey echo Install KEY :: !sKey! & %ospp% /inpkey:!sKey!
	if not defined sKey echo ERROR ### Key validation Failed
	
	echo:
	echo Integrate !pName! License [Backup method]
	if defined sKey set "PidKey=PidKey=!sKey!"
	"%root%\Integration\integrator" /I /License PRIDName=!pName!.16 !PidKey!
	
	REM *************
	REM  # M I S C #
	REM *************
	
	for /f "tokens=2 delims=_" %%g in ('reg query !key! /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
	set "ProductReleaseIds=!ProductReleaseIds:~6!"
	(echo !ProductReleaseIds! | >nul find /i "%pName%") || (
		>nul 2>nul reg add !key! /f /v ProductReleaseIds /t REG_SZ /d "!ProductReleaseIds!,%pName%"
	)
)

set "QueryClass=SoftwareLicensingProduct"
set "QueryTable=PartialProductKey, ID, Name"
set "QueryString=name like '%%!pFilter!%%'"

echo:
>nul 2>nul where wmic && (Call :UsingWMIC) & goto :END
>nul 2>nul where powershell && (Call :UsingPS) & goto :END
call :UsingVBS & goto :END

:end
exit /b

:USAGE
echo:
echo Usage :: Just Filter
echo Usage :: '%~nx0' Filter * *
echo:
echo Usage :: Filter + New Product Installaion
echo Usage :: '%~nx0' Filter ProductID Serial
echo:
echo Usage :: Filter + New Product Installaion + Clean Old product ID
echo Usage :: '%~nx0' Filter ProductID Serial -Clean
echo:
goto :eof

:Validate
set "sKey=#%~1#"
set "pattern=[0-z][0-z][0-z][0-z][0-z]"

if /i '%sKey:~0,1%' EQU '#' if /i '%sKey:~30,1%' EQU '#' if /i '%sKey:~30,2%' NEQ '##' (
	(echo "%sKey:~1,17%" | >nul findstr /r "%pattern%-%pattern%-%pattern%") && (echo "%sKey:~13,17%"  | >nul findstr /r "%pattern%-%pattern%-%pattern%") && set "sKey=%sKey:~1,-1%" && goto :eof
)
set "sKey="
goto :eof

:Clean_A
REM [1] cleanospp nethod
for %%$ in (msvcr100.dll, cleanospp.exe) do (	
	set "Dst=%windir%\%%$"
	set "Src=https://github.com/DarkDinosaurEx/OfficeRTool/blob/main/OfficeFixes/win_x32/%%$?raw=true"
	if not exist "%windir%\%%$" >nul 2>nul powershell -noprofile -executionpolicy bypass -command "Start-BitsTransfer -Source '!Src!' -Destination '!Dst!' -ea 1"
)
if exist "%windir%\cleanospp.exe" "%windir%\cleanospp.exe"
goto :eof

:Clean_M
REM [2] Manual nethod
>nul 2>nul reg add !key! /f /v ProductReleaseIds /t REG_SZ /d "%pName%"
for /f "tokens=1,8 delims= " %%g in ('%OSPP2% /dstatusall') do (
	if /i '%%g' EQU 'Last' (
		echo Remove  KEY :: XXXXX-XXXXX-XXXXX-XXXXX-%%h
		%OSPP% /unpkey:%%h
	)
)
goto :eof

:UsingWMIC
wmic path %QueryClass% where (%QueryString%) get %QueryTable%
goto :eof

:UsingPS
>"%temp%\script.ps1" echo Get-WmiObject -Query "select * from %QueryClass% where %QueryString%" ^| format-table -Property %QueryTable%
powershell -noprofile -executionpolicy bypass -file "%temp%\script.ps1"
goto :eof

:UsingVBS
 >"%temp%\script.vbs" echo on error resume next
>>"%temp%\script.vbs" echo strQuery = "Select " ^+ "%QueryTable%" + " from " ^+ "%QueryClass%" ^+ " where " ^+ "%QueryString%"
>>"%temp%\script.vbs" echo Set objArray= GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(strQuery,,48)
>>"%temp%\script.vbs" echo For each obj in objArray
>>"%temp%\script.vbs" echo 	result = ","
>>"%temp%\script.vbs" echo 	For each Prop in obj.Properties_
>>"%temp%\script.vbs" echo 		if NOT result = "," Then
>>"%temp%\script.vbs" echo 			result = result ^& Prop.Value ^& ","
>>"%temp%\script.vbs" echo 		End if
>>"%temp%\script.vbs" echo 		if result = "," Then
>>"%temp%\script.vbs" echo 			result = Prop.Value ^& ","
>>"%temp%\script.vbs" echo 		End if
>>"%temp%\script.vbs" echo 	Next
>>"%temp%\script.vbs" echo 	if NOT result = "," Then
>>"%temp%\script.vbs" echo 		WScript.Echo "%QueryTable%"
>>"%temp%\script.vbs" echo 		WScript.Echo result
>>"%temp%\script.vbs" echo end if
>>"%temp%\script.vbs" echo Next
cscript //Nologo "%temp%\script.vbs"
goto :eof