@cls
@echo off

set _days=SUN
set _time=19:00
set _frequency=weekly
::::::::::::::::::::::::::::::::::
set ScriptMode=WMI
set ConvertedMode=true
::::::::::::::::::::::::::::::::::
set LocalKms=true
::::::::::::::::::::::::::::::::::
set "SupportedbuildYear=2019, 2021"
::::::::::::::::::::::::::::::::::
set Invalid="Invalid"
set ProductError="ERROR:"
set ProductNotExist="No Instance(s) Available."
set ProductNotFound="Not found"
::::::::::::::::::::::::::::::::::
set Kms_Servers_List="%temp%\Kms_Servers_List.ini"
set Office_Keys_List="%temp%\Office_Keys_List.ini"
set Windows_Keys_List="%temp%\Windows_Keys_List.ini"
::::::::::::::::::::::::::::::::::
set logFile="%temp%\log"
set SecLogFile="%temp%\SecLogFile"
set "PrintNul=1>nul 2>nul"
set "PrintLog=1>>%LogFile% 2>>&1"
set "PrintCmd=1>>%LogFile% echo.&1>>%LogFile% echo"
set "PrintCmd_v2=1>>%LogFile% echo"
::::::::::::::::::::::::::::::::::
setlocal enabledelayedexpansion
title KMS Activation Tool ~ !ScriptMode! ~ Mode

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
)

:: basic validtion
call :cleanVariable
%PrintNul% del /q %logFile%
%PrintNul% del /q "%temp%\*.ini"

rem get rid of the not genuine banner solution by Windows_Addict
	
rem first NAG ~ check for IP address On start up
rem https://forums.mydigitallife.net/threads/kms_vl_all-smart-activation-script.79535/page-180#post-1659178

rem second NAG ~ check if ip address is from range of 0.0.0.0 to ?
rem HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing\LVUXRecords
rem https://forums.mydigitallife.net/threads/kms_vl_all-smart-activation-script.79535/page-237#post-1734148

rem How-to: Generate Random Numbers
rem https://ss64.com/nt/syntax-random.html

Set /a rand_A=192
Set /a rand_B=168
Set /a rand_C=(%RANDOM%*255/32768)+1
Set /a rand_D=(%RANDOM%*255/32768)+1

REM Set /a rand_A=(%RANDOM%*(255-192)/32768)+193
REM Set /a rand_B=(%RANDOM%*(255-168)/32768)+169

set "IP_ADDRESS=!rand_A!.!rand_B!.!rand_C!.!rand_D!"

:::: Run as Admin with native shell, any path, params, loop guard, minimal i/o, by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\" &call \"%%2\" %%3" & set "_= %*"
>nul fltmc || if "%f0%" neq "%~f0" ( cd.>"%tmp%\runas.Admin" & start "%~n0" /high "%tmp%\runas.Admin" "%~f0" "%_:"=""%" &exit /b )

set "WSH_Disabled=" & for %%$ in (HKCU, HKLM) do 2>nul reg query "%%$\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" | >nul find /i "0x0" && set "WSH_Disabled=***"
if defined WSH_Disabled for %%$ in (HKCU, HKLM) do %PrintNul% REG DELETE "%%$\SOFTWARE\Microsoft\Windows Script Host\Settings" /f /v Enabled
set "WSH_Disabled=" & for %%$ in (HKCU, HKLM) do 2>nul reg query "%%$\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" | >nul find /i "0x0" && set "WSH_Disabled=***"
if defined WSH_Disabled (
	cls
	echo.
	echo ERROR ### Windows script host is disabled
	echo.
	if not defined debugMode pause
	exit /b
)

cd /d "%~dp0"

if exist "Tools\KmsHelper.vbs" >nul 2>&1 copy /y "Tools\KmsHelper.vbs" %windir%
if not exist "%windir%\KmsHelper.vbs" (
	echo.
	echo Missing Critical Files [KmsHelper.vbs]
	echo.
	pause
	exit /b
)

if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set "7zr=tools\x32\7zr")
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set "7zr=tools\x64\7zr")
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set "7zr=tools\x64\7zr"
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set "7zr=tools\x64\7zr"

if defined LocalKms (

	if exist Tools\KmsService.7z (
		call :CmdWorker "!7zr!" x -r -y -o"%windir%" Tools\KmsService.7z  *.*
	)
	
	set "bin=A64.dll,x64.dll,x86.dll"
	REM set "x32FileList=vlmcsd.exe,SECOPatcher.dll,FakeClient.exe,libkms32.dll,libkms64.dll,WinDivert.dll,WinDivert.lib,WinDivert32.sys,WinDivert64.sys"
	REM set "x64FileList=vlmcsd.exe,SECOPatcher.dll,FakeClient.exe,libkms64.dll,WinDivert.dll,WinDivert.lib,WinDivert64.sys"
	
	for %%# in (!bin!) do if not exist "%windir%\bin\%%#" set "LocalKms="
	REM for %%# in (!x32FileList!) do if not exist "%windir%\x32\%%#" set "LocalKms="
	REM for %%# in (!x64FileList!) do if not exist "%windir%\x64\%%#" set "LocalKms="
	
	if not defined LocalKms (
		echo.
		echo Local Activation files Is Missing, Switch back to Online KMS
		timeout 5
		cls
	)
)

set "BuildNumber="
set "Windows_7_Or_Earlier="

call :query "buildnumber" "Win32_OperatingSystem"
for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set "BuildNumber=%%g"
set "BuildNumber=!BuildNumber: =!"

if !buildnumber! LSS 2600 (
	echo.
	echo ERROR ### System not supported
	echo.
	pause
	exit /b
)

if /i !BuildNumber! LEQ 7601 (
	set ScriptMode=WMI
	set "Windows_7_Or_Earlier=true"
	title KMS Activation Tool ~ !ScriptMode! ~ Mode
)

echo.
echo Kms Smart Activation Tool
echo Convert ^& Activate Smart Solution For Office / Windows Products
echo Support Activation for 			:: Office 2010 --^> late Office 2021
echo Support Convert For 			:: Office 2016 [MSI], 2016 --^> 2021 [C2R]
echo Support Convert / Activation for 	:: Windows Vista --^> Late Windows 11
echo.
echo ** Keep Origional Activated OEM / Retail / Mak Licences
echo ** Clean Duplicated Licences Of same Products, different year
echo ** Clean Unused Product Licences like :: 365, Home, Professional, Private
echo.
>nul timeout 3

if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set xBit=32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set xBit=64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set xBit=64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set xBit=64

call :CreateTask
>%Kms_Servers_List% 	call :Kms_Servers_List
>%Office_Keys_List% 	call :Office_Keys_List
>%Windows_Keys_List% 	call :Windows_Keys_List
call :LetsActivate

:end

rem Windows_Addict Not genuine fix
call :CleanRegistryKeys
call :CmdWorker reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"

%PrintNul% del /q "%temp%\*.ini"
setlocal disabledelayedexpansion
echo Log file saved at :: "%temp%\log"
echo please wait 6 seconds..
timeout /t 6 /NOBREAK
goto :eof

:cleanVariable
set vars=vbsActivation, cscript, slmgr, OSPP, OSPP_14, OSPP_15, OSPP_16, KmsServer, tokens, product, products, Period, args, tool, id, edition, oldGrace, key, root, guid, ProductYear, ProductIsActivated, tmpKey, ProductList, ConvertToMondo, ProductReleaseIds
for %%v in (%vars%) do (set %%v=)
goto :eof

:CreateTask
if exist "tools\tcping.exe" (
	echo Create automatic Task
	echo .....................
	%printCmd% copy /y "%~nx0" "%windir%\KmsActivationTool.cmd"
	(%PrintLog% copy /y "%~nx0" "%windir%\KmsActivationTool.cmd") && (
		%PrintLog% echo ******************************************
		call :CmdWorker copy /y tools\tcping.exe %windir%\tcping.exe
		if not exist %windir%\Licenses16 call :CmdWorker "!7zr!" x -r -y -o%windir% Tools\Licenses16.7z  *.*
		
		%PrintNul% pushd "%windir%\System32\WindowsPowershell\v1.0\" && (
			call :CmdWorker schtasks /create /SC %_frequency% /D %_days% /ST %_time% /RL HIGHEST /tn "KMS Activation Tool" /f /tr "powershell start %windir%\KmsActivationTool.cmd -Verb RunAs -WindowStyle hidden"
		) || (
			>"%windir%\KmsActivationTool.vbs" echo Set UAC ^= CreateObject^("Shell.Application"^).UAC.ShellExecute "%windir%\KmsActivationTool.cmd", "" , "", "runas", 0
			call :CmdWorker schtasks /create /SC %_frequency% /D %_days% /ST %_time% /RL HIGHEST /tn "KMS Activation Tool" /f /tr "%windir%\KmsActivationTool.vbs"
		)
		
		echo Done And Done
	) || (
		echo fail to create a Task
		%PrintLog% fail to create a Task
		%PrintLog% echo ******************************************
	)
	echo.
)
goto :eof

:LetsActivate
if defined LocalKms goto :vNext
set args=-4 -n 1 -g 1 -w 0.5 -i 0.5
set port=1688
set tool=%windir%\tcping.exe

if not exist !tool! (
	set "args="
	set "port="
	set "tool=ping"
)

:: Let find us some available Kms Servers
:: Let find us some available Kms Servers
:: Let find us some available Kms Servers

echo Start Activation Process
echo ........................
echo Look For Online Kms Servers
for /f "tokens=*" %%g in ('type %Kms_Servers_List%') do (
	echo Check if %%g:1688 is Online
	%printCmd% !tool! !args! %%g !port!
	%PrintLog% !tool! !args! %%g !port! && (
		%PrintLog% echo ******************************************
		set KmsServer=%%g
		goto :LetsActivate_
	)
	%PrintLog% echo ******************************************
)

echo.
echo ERROR ##### didnt found any available online kms server
echo.
goto :eof
:LetsActivate_
echo Winner Winner Chicken dinner
echo.

:vNext
if defined LocalKms (
	call :StartKMSActivation
	call :UpdateRegistryKeys !KMSHostIP! !KMSPort!
)

:: ugly check for X32 files
:: ugly check for X32 files
:: ugly check for X32 files

if exist "%windir%\System32\cscript.exe" 							set cscript="%windir%\System32\cscript.exe"
if exist "%windir%\System32\slmgr.vbs" 								set slmgr="%windir%\System32\slmgr.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\office14\OSPP.vbs" 	set OSPP_14="%ProgramFiles(x86)%\Microsoft Office\office14\OSPP.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\office15\OSPP.vbs" 	set OSPP_15="%ProgramFiles(x86)%\Microsoft Office\office15\OSPP.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs" 	set OSPP_16="%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\office15\OSPP.vbs" 	set OSPP="%ProgramFiles(x86)%\Microsoft Office\office15\OSPP.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs" 	set OSPP="%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.vbs"
if exist "%ProgramFiles(x86)%\Microsoft Office\root\Licenses16" 	set "licenceDir=%ProgramFiles(x86)%\Microsoft Office\root\Licenses16"
if exist "%ProgramFiles(x86)%\Microsoft Office\root" 				set "root=%ProgramFiles(x86)%\Microsoft Office\root"

(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun" /v PackageGUID) && (for /f "tokens=3" %%g in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\\WOW6432Node\Microsoft\Office\ClickToRun" /v PackageGUID') do set guid=%%g)

:: ugly check for x64 files
:: ugly check for x64 files
:: ugly check for x64 files

if exist "%windir%\SysWOW64\cscript.exe" 							set cscript="%windir%\SysWOW64\cscript.exe"
if exist "%windir%\SysWOW64\slmgr.vbs" 								set slmgr="%windir%\SysWOW64\slmgr.vbs"
if exist "%ProgramFiles%\Microsoft Office\office14\OSPP.vbs" 		set OSPP_14="%ProgramFiles%\Microsoft Office\office14\OSPP.vbs"
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.vbs" 		set OSPP_15="%ProgramFiles%\Microsoft Office\Office15\OSPP.vbs"
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs" 		set OSPP_16="%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs"
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.vbs" 		set OSPP="%ProgramFiles%\Microsoft Office\Office15\OSPP.vbs"
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs" 		set OSPP="%ProgramFiles%\Microsoft Office\Office16\OSPP.vbs"
if exist "%ProgramFiles%\Microsoft Office\root\Licenses16" 			set "licenceDir=%ProgramFiles%\Microsoft Office\root\Licenses16"
if exist "%ProgramFiles%\Microsoft Office\root" 					set "root=%ProgramFiles%\Microsoft Office\root"

(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds) && set Key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
(%PrintNul% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun" /v PackageGUID) && (for /f "tokens=3" %%g in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun" /v PackageGUID') do set guid=%%g)

:: Configure SPP 	Registry
:: Configure OSPP 	Registry

set OSPP_HKLM=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set OSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set XSPP_HKLM_X32=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_HKLM_X64=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
call :CleanRegistryKeys

:: on Win 7 and MSI version
:: we dont have licence files

if not defined licenceDir set licenceDir="%windir%\Licenses16"

call :WindowsHelper
call :OfficeHelper

if defined LocalKms (
	call :CleanRegistryKeys
	call :StopKMSActivation
)

echo.
goto :eof

:WindowsHelper

if not defined slmgr (
	echo ERROR ##### didnt found any Windows products / SLMGR.VBS IS Missing
	echo.
	goto :eof
)

:: Check if we got Retail Or Mak Licences
call :Search_VL_Products Windows
set VL_Product_Not_Found=
(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %Invalid%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true

if defined VL_Product_Not_Found (

	:: check if product Is Retail activated
	call :Check_Activation Windows RETAIL
	if defined ProductIsActivated goto :eof
	
	:: check if product Is MAK activated
	call :Check_Activation Windows MAK
	if defined ProductIsActivated goto :eof
	
	:: check if product Is OEM activated
	call :Check_Activation Windows OEM
	if defined ProductIsActivated goto :eof
	
	:: if not. convert to VL.
	call :Windows_Licence_Worker
	
	if not defined serial goto :eof

	:: make sure we got Volume Licence
	set VL_Product_Not_Found=
	call :InstallProductKey !serial!
	call :Search_VL_Products Windows
	(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %invalid%>nul) && set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true	
	(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true

	if defined VL_Product_Not_Found (
		echo ERROR ##### didnt found any windows volume products
		goto :eof
	)
)

for /f "tokens=1,2,3,4,5,6,7,8 skip=3 delims=," %%g in ('type "%temp%\result"') do (
	
	set "editionName=%%l"
	set /a "Period=%%h/60/24"
	
	echo %%k !editionName:~1!
	echo ....................................
	echo Partial Product Key	= %%m
	echo License / Genuine 	= %%j / %%g
	echo Period Remaining 	= !Period! Days
	echo Id 			= %%i
	
	if %%h GTR 259200 (
		echo.
		echo ^=^=^= Windows is KMS38 activated ^=^=^=
		echo.
	)
	
	if /i '!ScriptMode!' EQU 'WMI' (
		if %%h LEQ 259200 (
			if defined LocalKms (
				call :Wmi_Activation %%i %%h !KMSHostIP! !KMSPort! windows
			) else (
				call :Wmi_Activation %%i %%h !KmsServer! 1688 windows
			)
		)
	)
	
	if /i '!ScriptMode!' EQU 'VBS' (
		if %%h LEQ 259200 (
			if defined LocalKms (
				call :Wmi_Activation %%i %%h !KMSHostIP! !KMSPort! windows
			) else (
				call :Vbs_Activation !KmsServer! 1688 windows
			)
		)
	)
)

goto :eof

:OfficeHelper

set xArgs=
set xArgs=%*
set 14.X_Mode=

if not defined OSPP_14 (
	if not defined OSPP_15 (
		if not defined OSPP_16 (
			echo ERROR ##### didnt found any Office products / OSPP.VBS IS Missing
			echo. & goto :eof
		)
	)
)

if defined xArgs (
	if /i '!xArgs!' EQU '14.X' (
		set vbsActivation=
		set 14.X_Mode=true
	)
)

:: if not. convert to VL.
if not defined 14.X_Mode (
	if defined ConvertedMode (
		call :Office_Licence_Worker
	)
)

:: lets try again
set VL_Product_Not_Found=
if defined 14.X_Mode 		call :Search_14.X_VL_Products
if not defined 14.X_Mode 	call :Search_VL_Products Office

(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %invalid%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true

if defined VL_Product_Not_Found (

	REM after check for Office 14.x
	if defined 14.X_Mode (
		echo ERROR ##### didnt found any Office 14.X volume products
		goto :eof
	)
	
	REM before check for Office 14.x
	REM Office 14.x Installed
	if defined OSPP_14 (
		echo ERROR ##### didnt found any Office 15.X 16.X volume products
		echo.
		call :OfficeHelper 14.X
		goto :eof
	)
	
	REM before check for Office 14.x
	REM Office 14.x Not Installed
	if not defined OSPP_14 (
		echo ERROR ##### didnt found any Office 14.X 15.X 16.X volume products
		goto :eof
	)
	
	echo ERROR ##### Wtf happen now ^?^?^?
	goto :eof
)

for /f "tokens=1,2,3,4,5,6,7,8 skip=3 delims=," %%g in ('type "%temp%\result"') do (
	
	set "editionName=%%l"
	set /a "Period=%%h/60/24"
	
	echo !editionName:~1, -8!
	echo ....................................
	echo Partial Product Key	= %%m
	echo License / Genuine 	= %%j / %%g
	echo Period Remaining 	= !Period! Days
	echo Id 			= %%i
		
	rem Activation Process
	rem 2 Options
	
	if /i '!ScriptMode!' EQU 'WMI' (
		if defined LocalKms (
			call :Wmi_Activation %%i %%h !KMSHostIP! !KMSPort! office
		) else (
			call :Wmi_Activation %%i %%h !KmsServer! 1688 office
		)
	)
	
	if /i '!ScriptMode!' EQU 'VBS' (
		if defined LocalKms (
			call :Wmi_Activation %%i %%h !KMSHostIP! !KMSPort! office
		) else (
			call :Vbs_Activation !KmsServer! 1688 office
		)
	)
)

if not defined 14.X_Mode (
	if defined OSPP_14 (
		call :OfficeHelper 14.X
		goto :eof
	)
	
	echo Search for 14.X Products
	echo ........................
	echo --- 404 not found
	goto :eof
)

goto :eof

:: Functions to find any retail products
:: and convert them to VL

:: Windows_Licence_Worker
:: Find Id and match it to VL Serial

:: Office_Licence_Worker
:: Find any Office Products and convert them to VL
:: without check if allready VL or not
:: it take time we dont have

:Windows_Licence_Worker

:: We need to match the edition
:: to specific Volume Kms Edition
:: because somtimes we can get somthing like this
:: all of this is volume But Edition IS Pro...
:: and we need Volume id of Pro...

:: 2de67392-b7a7-462a-b1ca-108dd189f588
:: 3f1afc82-f8ac-4f6c-8005-1d233e606eee
:: 73111121-5638-40f6-bc11-f1d7b0d64300
:: 82bbc092-bc50-4e16-8e18-b74fc486aec3
:: e0c42288-980c-4788-a014-c080d2e1926e
:: e4db50ea-bda1-4566-b047-0ca50abc6f07
:: ec868e65-fadf-4759-b23e-93fe37f2cc29

set serial=
set windowsID=
set VL_Product_Not_Found=

for /f "skip=7 tokens=2 delims=:" %%g in ('dism /English /Online /Get-CurrentEdition') do set edition=%%g
set wmiSearch=Name like '%%%%!edition!%%%%' and Description like '%%%%VOLUME_KMSCLIENT%%%%' and ApplicationId like '%%%%55c92734-d682-4d71-983e-d6ec3f16059f%%%%'
call :query "ID" "SoftwareLicensingProduct" "!wmiSearch!"

(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %invalid%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true
if defined VL_Product_Not_Found (
	echo ERROR ##### Couldn't find Any windows Supported ID
	goto :eof
)

for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do (if not defined windowsID set "windowsID=%%g")
if not defined windowsID goto :eof

for %%A in (
	b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
	11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
	6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
	4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
	2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do ( if /i '!windowsID!' equ '%%A' (
	set "windowsID="
	exit /b ))
for /f "tokens=1,2 delims=*" %%r in ('type %Windows_Keys_List%') do (if /i '%%r' equ '!windowsID!' set "serial=%%s")
goto :eof

:Office_Licence_Worker

:: check if we have office products at all
call :GetProductYear

:: no installed office ...
if not defined ProductYear goto :eof

:: there is office products
if defined Officec2r (
	set ProductReleaseIds=
	for /f "tokens=2 delims=_" %%g in ('reg query %key% /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
	set ProductReleaseIds=!ProductReleaseIds:~6!
)

if not defined OfficeMsi16 (
	if not defined ProductReleaseIds (
		goto :eof
	)
)

set ProductList=365,HOME,Professional,Private
for %%x in (!ProductList!) do (

	set SelectedX=
	if defined Officec2r (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
		)
	)	
	if defined OfficeMsi16 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined SelectedX (
	
		set MoveToNext=
		set ProductReleaseIds_=
		
		if /i '%%x' EQU 'HOME' (
			set ProductReleaseIds_=!ProductReleaseIds: =!
			for %%g in (!ProductReleaseIds_!) do (
				echo %%g | >nul find /i "home" && (
					echo %%g | >nul find /i "365" && set MoveToNext=true
				)
			)
		)
		
		if not defined MoveToNext (
			set Office_Product_Not_Found=true
		
			:: check if product Is Retail activated
			call :Check_Activation %%x RETAIL
			if defined ProductIsActivated set Office_Product_Not_Found=

			:: check if product Is MAK activated
			call :Check_Activation %%x MAK
			if defined ProductIsActivated set Office_Product_Not_Found=
			
			if defined Office_Product_Not_Found (
				set ConvertToMondo=true
				call :Remove_Office_Products !ProductYear! %%x
			)
		)
	)
)

if defined ConvertToMondo (
	
	set tYear=2016
	set ProductYear=2016
	set tProduct=Mondo
	
	set VL_Product_Not_Found=
	call :Search_Office_VL_Products !tYear! !tProduct!
	(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %invalid%>nul) && set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true
	(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true

	if defined VL_Product_Not_Found (
	
		:: check if product Is Retail activated
		call :Check_Activation !tProduct! RETAIL
		if defined ProductIsActivated set VL_Product_Not_Found=

		:: check if product Is MAK activated
		call :Check_Activation !tProduct! MAK
		if defined ProductIsActivated set VL_Product_Not_Found=

		if defined VL_Product_Not_Found (
		
			call :integrate !tProduct! !ProductYear!
			for /f %%y in ('dir /b "%licenceDir%\!tProduct!VL_KMS*.xrm-ms"') do call :inslic %%y

			:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
			for /f "tokens=1,2,3 delims=*" %%p in ('type %Office_Keys_List%') do (
				if /i '%%q' EQU '!tYear!' (
					echo %%p | find /i "!tProduct!">nul && call :inpkey %%r
				)
			)
		)
	)
	
	set ProductList=publisher,ProjectPro,ProjectStd,VisioStd,VisioPro
	goto :Office_Worker_begin
)

set ProductList=proplus,Standard,mondo,word,excel,powerpoint,Skype,access,outlook,publisher,ProjectPro,ProjectStd,VisioStd,VisioPro,OneNote

:Office_Worker_begin

for %%x in (!ProductList!) do (

	set SelectedX=
	if defined OfficeMsi16 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
			
			if not defined ConvertToMondo (
				for %%m in (!ProductReleaseIds!) do (
					echo %%m | >nul find /i "%%x" && (
						set ProductYear=
						echo %%m | >nul find /i "2021" && 	set ProductYear=2021
						echo %%m | >nul find /i "2019" && 	set ProductYear=2019
						if not defined ProductYear 		  	set ProductYear=2016
						if /i "%%x" EQU "OneNote" 			set ProductYear=2016
					)
				)
			)
		)
	)
	if defined SelectedX (
	
		set VL_Product_Not_Found=
		call :Search_Office_VL_Products !ProductYear! %%x
		(type "%temp%\result"|find /i ",">nul) || set VL_Product_Not_Found=true
		(type "%temp%\result"|find /i %invalid%>nul) && set VL_Product_Not_Found=true
		(type "%temp%\result"|find /i %ProductError%>nul) && set VL_Product_Not_Found=true
		(type "%temp%\result"|find /i %ProductNotExist%>nul) && set VL_Product_Not_Found=true
		(type "%temp%\result"|find /i %ProductNotFound%>nul) && set VL_Product_Not_Found=true
		
		if defined VL_Product_Not_Found (
		
			:: check if product Is Retail activated
			call :Check_Activation %%x RETAIL
			if defined ProductIsActivated set VL_Product_Not_Found=

			:: check if product Is MAK activated
			call :Check_Activation %%x MAK
			if defined ProductIsActivated set VL_Product_Not_Found=
			
			if defined VL_Product_Not_Found (
			
				REM call :Remove_Office_Products !ProductYear! %%x
				for %%$ in (2016, 2019, 2021) do call :Remove_Office_Products %%$ %%x
				
				set "year="
				if /i '!ProductYear!' NEQ '2016' set "year=!ProductYear!"
				
				call :integrate %%x !ProductYear!
				for /f %%y in ('dir /b "%licenceDir%\%%x!year!VL_KMS*.xrm-ms"') do call :inslic %%y
				
				:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
				for /f "tokens=1,2,3 delims=*" %%p in ('type %Office_Keys_List%') do (
					if /i '%%q' EQU '!ProductYear!' (
						echo %%p | find /i "%%x">nul && call :inpkey %%r
					)
				)
			)
		)
	)
)

goto :eof

:: General 	Function
:: Windows 	Function
:: Office 	Function

:CmdWorker

set pStart=%time: =0%
%PrintNul% del /q %SecLogFile%
1>>%SecLogFile% 2>>&1 %*
set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
%printCmd_v2% %*
if exist %SecLogFile% %PrintLog% type %SecLogFile%
%PrintLog% echo ******************************************

goto :eof

:Check_Activation

:: Check for permenent activation
:: first args can be office or windows
:: second args can be retail or mak or oem

set product=%1
set licenceType=%2
set ProductIsActivated=true
set wmiSearch=Name like '%%%%!product!%%%%' and Description like '%%%%!licenceType!%%%%' and PARTIALPRODUCTKEY IS NOT NULL and GenuineStatus = 0 and LicenseStatus = 1

set LicensingProductClass=SoftwareLicensingProduct
echo %1 | >nul find /i "Office" && (if defined Windows_7_Or_Earlier set LicensingProductClass=OfficeSoftwareProtectionProduct)

call :query "name" "!LicensingProductClass!" "!wmiSearch!"
(type "%temp%\result" | find /i "," >nul) || set ProductIsActivated=
(type "%temp%\result" | find /i %invalid% >nul) && set ProductIsActivated=
(type "%temp%\result" | find /i %ProductError% >nul) && set ProductIsActivated=
(type "%temp%\result" | find /i %ProductNotExist% >nul) && set ProductIsActivated=
(type "%temp%\result" | find /i %ProductNotFound% >nul) && set ProductIsActivated=

if defined ProductIsActivated (
	if /i "%product%" EQU "Windows" echo ^=^=^= !product! is !licenceType! activated ^=^=^=
	if /i "%product%" NEQ "Windows" echo ^=^=^= Office !product! is !licenceType! activated ^=^=^=
	echo.
)
goto :eof

:Search_VL_Products

set ApplicationId=
set LicensingProductClass=SoftwareLicensingProduct

echo %1 | >nul find /i "Office" && (
	set ApplicationId=0ff1ce15-a989-479d-af46-f275c6370663
	if defined Windows_7_Or_Earlier set LicensingProductClass=OfficeSoftwareProtectionProduct
)
echo %1 | >nul find /i "windows" && set ApplicationId=55c92734-d682-4d71-983e-d6ec3f16059f

set info=ID, LicenseStatus, PartialProductKey, GenuineStatus, Name, GracePeriodRemaining
set wmiSearch=ApplicationId like '%%%%!ApplicationId!%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null

call :Query "!info!" "!LicensingProductClass!" "!wmiSearch!"
goto :eof

:Search_14.X_VL_Products

set ApplicationId=59a52881-a989-479d-af46-f275c6370663
set LicensingProductClass=OfficeSoftwareProtectionProduct
set info=ID, LicenseStatus, PartialProductKey, GenuineStatus, Name, GracePeriodRemaining
set wmiSearch=ApplicationId like '%%%%!ApplicationId!%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null

call :Query "!info!" "!LicensingProductClass!" "!wmiSearch!"
goto :eof

:Search_Office_VL_Products

set T_Year=%1
Set T_Name=%2
set LicensingProductClass=SoftwareLicensingProduct
set ApplicationId=0ff1ce15-a989-479d-af46-f275c6370663
if defined Windows_7_Or_Earlier set LicensingProductClass=OfficeSoftwareProtectionProduct

set wmiSearch=ApplicationId like '%%%%!ApplicationId!%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null and Name like '%%%%office !T_year:~2!%%%%' And Name like '%%%%!T_Name!%%%%'
call :Query "Name" "!LicensingProductClass!" "!wmiSearch!"

goto :eof

:Remove_Office_Products
set T_Year=%1
Set T_Name=%2
set LicensingProductClass=SoftwareLicensingProduct
if defined Windows_7_Or_Earlier set LicensingProductClass=OfficeSoftwareProtectionProduct

set info=ID, LicenseStatus, PartialProductKey, GenuineStatus, Name, GracePeriodRemaining
set wmiSearch=PartialProductKey is not null and Name like '%%%%office !T_year:~2!%%%%' And Name like '%%%%!T_Name!%%%%'
call :UninstallProductKey "!LicensingProductClass!" "!wmiSearch!"
goto :eof

:Wmi_Activation

:: Wmi based activation
:: using SoftwareLicensingProduct class for office
:: using SoftwareLicensingService class for windows
:: using OfficeSoftwareProtectionProduct class for specific win7 case

set Id=%1
set Ks=%3
set Kp=%4
set Km=%5

set SPP_ACT_CLASS=
set SPP_KMS_Class=
set SPP_KMS_Where=

(echo "%Km%" | >nul find /i "windows") && (

	rem Windows activation using 2 classes
	rem SoftwareLicensingService for KMS settings
	rem SoftwareLicensingProduct for activation
	
	set subKey=55c92734-d682-4d71-983e-d6ec3f16059f
	set SPP_KMS_Class=SoftwareLicensingService
	set SPP_KMS_Where=version is not null
	set SPP_ACT_CLASS=SoftwareLicensingProduct
)

(echo "%Km%" | >nul find /i "office") && (
	
	set subKey=0ff1ce15-a989-479d-af46-f275c6370663
	
	if defined Windows_7_Or_Earlier (
	
		rem Office in windows Windows 7 and less use 2 classes
		rem OfficeSoftwareProtectionService for KMS settings 
		rem OfficeSoftwareProtectionProduct for activation
		
		set SPP_KMS_Class=OfficeSoftwareProtectionService
		set SPP_KMS_Where=version is not null
		set SPP_ACT_CLASS=OfficeSoftwareProtectionProduct
	)
	
	if defined 14.X_Mode (
		
		rem Office 2010 Classes
		rem OfficeSoftwareProtectionService for KMS settings 
		rem OfficeSoftwareProtectionProduct for activation
		
		set subKey=59a52881-a989-479d-af46-f275c6370663
		set SPP_KMS_Class=OfficeSoftwareProtectionService
		set SPP_KMS_Where=version is not null
		set SPP_ACT_CLASS=OfficeSoftwareProtectionProduct
	)
)

rem case of office And Not Win 7
rem case of office And Not Win 7
rem case of office And Not Win 7

if not defined SPP_ACT_CLASS set SPP_ACT_CLASS=SoftwareLicensingProduct
set Product_Licensing_Class=!SPP_ACT_CLASS!
set Product_Licensing_Where=Id like '%%%%!Id!%%%%'
if not defined SPP_KMS_Class (
	set SPP_KMS_Class=!Product_Licensing_Class!
	set SPP_KMS_Where=!Product_Licensing_Where!
)

:: lets go to work
:: lets go to work
:: lets go to work

call :UpdateRegistryKeys %Ks% %Kp%

:: activation always using SoftwareLicensingProduct
echo.
echo ^+^+^+ Activating ^+^+^+ && echo ...................
call :ACTIVATE !Product_Licensing_Class! !Id!

REM :: compare old time to new time
call :Query "GracePeriodRemaining" "!Product_Licensing_Class!" "!Product_Licensing_Where!"
for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set "GracePeriod=%%g"
set GracePeriod=!GracePeriod: =!
echo Old Grace               = %2
echo New Grace               = !GracePeriod!
if not defined lastErr echo Status                  = Unknown
if defined lastErr if /i '!lastErr!' NEQ '0' (echo Status                  = Failed ^(Error 0x%lastErr%^)) else (echo Status                  = Succeeded ^(Error 0x%lastErr%^))
(echo "%Km%" | >nul find /i "windows") 				&& (echo Product Class           = SoftwareLicensingService) 		&& (echo Product Class           = SoftwareLicensingProduct) && goto :skip_
echo !SPP_KMS_Class! 			| >nul find /i "Office" 	&& (echo Product Class           = OfficeSoftwareProtectionService) || (echo Product Class           = SoftwareLicensingProduct)
echo !Product_Licensing_Class! 	| >nul find /i "Office" 	&& (echo Licensing Class         = OfficeSoftwareProtectionProduct) || (echo Licensing Class         = SoftwareLicensingProduct)

:skip_
echo.
goto :eof

:Vbs_Activation

:: VBS based activation
:: using OSPP for office
:: using slmgr for windows

set Ks=%1
set Kp=%2
set Km=%3

set Id=
set subKey=

echo "%Km%" | >nul find /i "windows" && (

	call :CmdWorker %cscript% %slmgr% /skms %Ks%:%Kp%
	call :CmdWorker %cscript% %slmgr% /ato
	(type %SecLogFile% | find /i "Product activated successfully.">nul) && (echo Status 			= Succeeded) || (echo Status 			= Failed)
	call :CmdWorker %cscript% %slmgr% /ckms
	call :CmdWorker %cscript% %slmgr% /cdns
	call :CmdWorker %cscript% %slmgr% /ckhc
)

echo "%Km%" | >nul find /i "office" && (

	REM make sure it run Once !
	if defined vbsActivation (
		echo Status 			= Same
		echo.
		goto :eof
	)
	set vbsActivation=true
	if defined 14.X_Mode set "OSPP=!OSPP_14!"
	
	call :CmdWorker %cscript% !OSPP! /sethst:%Ks%
	call :CmdWorker %cscript% !OSPP! /setprt:%Kp%
	call :CmdWorker %cscript% !OSPP! /act
	(type %SecLogFile% | find /i "Product activation successful">nul) && (echo Status 			= Succeeded) || (echo Status 			= Failed)
	call :CmdWorker %cscript% !OSPP! /remhst
)

echo.
goto :eof

:CleanRegistryKeys

rem OSPP.VBS Nethood
rem OSPP.VBS Nethood
rem OSPP.VBS Nethood

call :CmdWorker reg delete "%OSPP_USER%" /f /v KeyManagementServiceName
call :CmdWorker reg delete "%OSPP_USER%" /f /v KeyManagementServicePort
call :CmdWorker reg delete "%OSPP_USER%" /f /v DisableDnsPublishing
call :CmdWorker reg delete "%OSPP_USER%" /f /v DisableKeyManagementServiceHostCaching

call :CmdWorker reg delete "%OSPP_HKLM%" /f /v KeyManagementServiceName
call :CmdWorker reg delete "%OSPP_HKLM%" /f /v KeyManagementServicePort
call :CmdWorker reg delete "%OSPP_HKLM%" /f /v DisableDnsPublishing
call :CmdWorker reg delete "%OSPP_HKLM%" /f /v DisableKeyManagementServiceHostCaching

rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood

call :CmdWorker reg delete "%XSPP_USER%" /f /v KeyManagementServiceName
call :CmdWorker reg delete "%XSPP_USER%" /f /v KeyManagementServicePort
call :CmdWorker reg delete "%XSPP_USER%" /f /v DisableDnsPublishing
call :CmdWorker reg delete "%XSPP_USER%" /f /v DisableKeyManagementServiceHostCaching

call :CmdWorker reg delete "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName
call :CmdWorker reg delete "%XSPP_HKLM_X32%" /f /v KeyManagementServicePort
call :CmdWorker reg delete "%XSPP_HKLM_X32%" /f /v DisableDnsPublishing
call :CmdWorker reg delete "%XSPP_HKLM_X32%" /f /v DisableKeyManagementServiceHostCaching

call :CmdWorker reg delete "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName
call :CmdWorker reg delete "%XSPP_HKLM_X64%" /f /v KeyManagementServicePort
call :CmdWorker reg delete "%XSPP_HKLM_X64%" /f /v DisableDnsPublishing
call :CmdWorker reg delete "%XSPP_HKLM_X64%" /f /v DisableKeyManagementServiceHostCaching

rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY

for %%# in (55c92734-d682-4d71-983e-d6ec3f16059f, 0ff1ce15-a989-479d-af46-f275c6370663, 59a52881-a989-479d-af46-f275c6370663) do (
	call :CmdWorker reg delete "%XSPP_USER%\%%#" /f
	call :CmdWorker reg delete "%XSPP_HKLM_X32%\%%#" /f
	call :CmdWorker reg delete "%XSPP_HKLM_X64%\%%#" /f
)
goto :eof

:updateRegistryKeys

rem OSPP.VBS Nethood
rem OSPP.VBS Nethood
rem OSPP.VBS Nethood

call :CmdWorker reg add "%OSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%OSPP_USER%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
call :CmdWorker reg add "%OSPP_USER%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
call :CmdWorker reg add "%OSPP_USER%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

call :CmdWorker reg add "%OSPP_HKLM%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%OSPP_HKLM%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
call :CmdWorker reg add "%OSPP_HKLM%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
call :CmdWorker reg add "%OSPP_HKLM%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood

call :CmdWorker reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_USER%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
call :CmdWorker reg add "%XSPP_USER%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
call :CmdWorker reg add "%XSPP_USER%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY

call :CmdWorker reg add "%XSPP_USER%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_USER%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

call :CmdWorker reg add "%XSPP_HKLM_X32%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_HKLM_X32%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

call :CmdWorker reg add "%XSPP_HKLM_X64%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
call :CmdWorker reg add "%XSPP_HKLM_X64%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

goto :eof

:GetProductYear

set Officec2r=
Set OfficeMsi16=

if not defined key (
	(reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16") && (
		Set OfficeMsi16=true
		call :update_Year 2016
	)
	(reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16") && (
		Set OfficeMsi16=true
		call :update_Year 2016
	)
	goto :eof
)

set ProductReleaseIds=
for /f "tokens=2 delims=_" %%g in ('reg query %key% /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
set ProductReleaseIds=!ProductReleaseIds:~6!
if not defined ProductReleaseIds goto :eof

set Officec2r=true
for %%x in (%SupportedbuildYear%) do (
	(echo !ProductReleaseIds! | >nul find /i "%%x") && (
		call :update_Year %%x
		goto :eof
	)
)

call :update_Year 2016
goto :eof

:update_Year
set ProductYear=%*
goto :eof

:InstallProductKey
call :CmdWorker %cscript% %slmgr% /ipk "%*"
goto :eof

:inslic
call :CmdWorker %cscript% !OSPP! /inslic:"%licenceDir%\%*"
goto :eof

:inpkey
call :CmdWorker %cscript% !OSPP! /inpkey:%*
goto :eof

:unpkey
set tmpKey=%*
set LicensingProductClass=SoftwareLicensingProduct
if defined Windows_7_Or_Earlier set LicensingProductClass=OfficeSoftwareProtectionProduct
if /i '!ScriptMode!' EQU 'VBS' call :CmdWorker %cscript% !OSPP! /unpkey:%*
if /i '!ScriptMode!' EQU 'WMI' call :UninstallProductKey "!LicensingProductClass!" "partialProductKey like '%%%%!tmpKey!%%%%'"
goto :eof

:integrate
set nYear=%2
set nProduct=%1
if not defined root goto :eof

(echo '%nYear%' | find /i "2016">nul) && (set nProduct=%nProduct%Volume.16) || (set nProduct=%nProduct%%nYear%Volume.16)
call :CmdWorker "%root%\Integration\integrator" /I /License PRIDName=!nProduct! PackageGUID="%guid%" /Global /C2R PackageRoot="%root%"
goto :eof

:CalcDiff
rem "Lean and Mean" TIMER with Regional format, 24h and mixed input support
rem https://stackoverflow.com/a/43968482

set timer_set=%1
set timer_end=%2
for /f "tokens=1-6 delims=0123456789" %%i in ("%timer_end%%timer_set%") do (set CE=%%i&set DE=%%k&set CS=%%l&set DS=%%n)
set "TE=!timer_end:%DE%=%%100)*100+1!"     & set "TS=!timer_set:%DS%=%%100)*100+1!"
set/A "T=((((10!TE:%CE%=%%100)*60+1!%%100)-((((10!TS:%CS%=%%100)*60+1!%%100)" & set/A "T=!T:-=8640000-!"
set/A "cc=T%%100+100,T/=100,ss=T%%60+100,T/=60,mm=T%%60+100,hh=T/60+100"
set "value=!hh:~1!%CE%!mm:~1!%CE%!ss:~1!%DE%!cc:~1!" & if "%~2"=="" echo/!value!
set "_tdiff=%value%" & set "timer_set=" & goto :eof
goto :eof

:: secret list
:: secret list
:: secret list

:Kms_Servers_List
echo kms.digiboy.ir
echo hq1.chinancce.com
echo kms.cnlic.com
echo kms.chinancce.com
echo kms.ddns.net
echo franklv.ddns.net
echo k.zpale.com
echo m.zpale.com
echo mvg.zpale.com
echo kms.shuax.com
echo kensol263.imwork.net
echo annychen.pw
echo heu168.6655.la
echo xykz.f3322.org
echo kms789.com
echo dimanyakms.sytes.net
echo kms.03k.org
echo kms.lotro.cc
echo kms.didichuxing.com
echo zh.us.to
echo kms.aglc.cckms.aglc.cc
echo kms.xspace.in
echo winkms.tk
echo kms.srv.crsoo.com
echo kms.loli.beer
echo kms8.MSGuides.com
echo kms9.MSGuides.com
echo kms.zhuxiaole.org
echo kms.lolico.moe
echo kms.moeclub.org
goto :eof

:Office_Keys_List
echo Word*2016*WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
echo word*2019*PBX3G-NWMT6-Q7XBW-PYJGG-WXD33
echo word*2021*TN8H9-M34D3-Y64V9-TR72V-X79KV
echo Excel*2016*9C2PK-NWTVB-JMPW8-BFT28-7FTBF
echo Excel*2019*TMJWT-YYNMB-3BKTF-644FC-RVXBD
echo Excel*2021*NWG3X-87C9K-TC7YY-BC2G7-G6RVC
echo PowerPoint*2016*J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
echo PowerPoint*2019*RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
echo PowerPoint*2021*TY7XF-NFRBR-KJ44C-G83KF-GX27K
echo ProPlus*2016*XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo ProPlus*2019*NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo ProPlus*2021*FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
echo VisioPro*2016*PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
echo VisioPro*2019*9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
echo VisioPro*2021*KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
echo VisioStd*2016*7WHWN-4T7MP-G96JF-G33KR-W8GF4
echo VisioStd*2019*7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
echo VisioStd*2021*MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
echo Publisher*2016*F47MM-N3XJP-TQXJ9-BP99D-8K837
echo Publisher*2019*G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
echo Publisher*2021*2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ
echo ProjectPro*2016*YG9NW-3K39V-2T3HJ-93F3Q-G83KT
echo ProjectPro*2019*B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
echo ProjectPro*2021*FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
echo ProjectStd*2016*GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
echo ProjectStd*2019*C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
echo ProjectStd*2021*J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
echo SkypeforBusiness*2016*869NQ-FJ69K-466HW-QYCP2-DDBV6
echo SkypeforBusiness*2019*NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
echo SkypeforBusiness*2021*HWCXN-K3WBT-WJBKY-R8BD9-XK29P
echo Standard*2016*JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo Standard*2019*6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo Standard*2021*KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
echo mondo*2016*HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
echo outlook*2016*R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
echo outlook*2019*7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
echo outlook*2021*C9FM6-3N72F-HFJXB-TM3V9-T86R9
echo access*2016*GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
echo access*2019*9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT
echo access*2021*WM8YG-YNGDD-4JHDC-PG3F4-FC4T4
echo OneNote*2016*DR92N-9HTF2-97XKM-XW2WJ-XW3J6
goto :eof

:Windows_Keys_List
echo ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03*WX4NM-KYWYW-QJJR4-XV3QB-6VM33
echo 9bd77860-9b31-4b7b-96ad-2564017315bf*VDYBN-27WPP-V4HQT-9VMD4-VMK7H
echo 8c8f0ad3-9a43-4e05-b840-93b8d1475cbc*6N379-GGTMK-23C6M-XVVTC-CKFRQ
echo 58e97c99-f377-4ef1-81d5-4ad5522b5fd8*TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
echo 7b9e1751-a8da-4f75-9560-5fadfe3d8e38*3KHY7-WNT83-DGQKR-F7HPR-844BM
echo cd918a57-a41b-4c82-8dce-1a538e221a83*7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
echo a9107544-f4a0-4053-a96a-1479abdef912*PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
echo 2de67392-b7a7-462a-b1ca-108dd189f588*W269N-WFGWX-YVC9B-4J6C9-T83GX
echo a80b5abf-76ad-428b-b05d-a47d2dffeebf*MH37W-N47XK-V7XM9-C7227-GCQG9
echo 3f1afc82-f8ac-4f6c-8005-1d233e606eee*6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
echo 5300b18c-2e33-4dc2-8291-47ffcec746dd*YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
echo 82bbc092-bc50-4e16-8e18-b74fc486aec3*NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
echo 4b1571d3-bafb-4b40-8087-a961be2caf65*9FNHH-K3HBT-3W4TD-6383H-6XYWF
echo e0c42288-980c-4788-a014-c080d2e1926e*NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
echo 3c102355-d027-42c6-ad23-2e7ef8a02585*2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
echo 73111121-5638-40f6-bc11-f1d7b0d64300*NPPR9-FWDCX-D2C8J-H872K-2YT43
echo e272e3e2-732f-4c65-a8f0-484747d0d947*DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
echo e0b2d383-d112-413f-8a80-97f373a5820c*YYVX9-NTFWV-6MDM3-9PT4T-4M68B
echo e38454fb-41a4-4f59-a5dc-25080e354730*44RPN-FTY23-9VTTB-MP9BX-T84FV
echo 7b51a46c-0c04-4e8f-9af4-8496cca90d5e*WNMTR-4C88C-JK8YV-HQ7T2-76DF9
echo 87b838b7-41b6-4590-8318-5797951d8529*2F77B-TNFGY-69QQF-B8YKP-D69TJ
echo 2d5a5a60-3040-48bf-beb0-fcd770c20ce0*DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
echo 9f776d83-7156-45b2-8a5c-359b9c9f22a3*QFFDN-GRT3P-VKWWX-X7T3R-8B639
echo 32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee*M7XTQ-FN8P6-TTKYV-9D4CC-J462D
echo 7103a333-b8c8-49cc-93ce-d37c09687f92*92NFX-8DJQP-P6BBQ-THF9C-7CG2H
echo e4db50ea-bda1-4566-b047-0ca50abc6f07*7NBT4-WGBQX-MP4H7-QXFF8-YP3KX
echo ec868e65-fadf-4759-b23e-93fe37f2cc29*CPWHC-NT2C7-VYW78-DHDB2-PG3GK
echo 0df4f814-3f57-4b8b-9a9d-fddadcd69fac*NBTWJ-3DR69-3C4V8-C26MC-GQ9M6
echo 034d3cbb-5d4b-4245-b3f8-f84571314078*WVDHN-86M7X-466P6-VHXV7-YY726
echo de32eafd-aaee-4662-9444-c1befb41bde2*N69G4-B89J2-4G8F4-WWYCC-J464C
echo 34e1ae55-27f8-4950-8877-7a03be5fb181*WMDGN-G9PQG-XVVXX-R3X43-63DFG
echo 73e3957c-fc0c-400d-9184-5f7b6f2eb409*N2KJX-J94YW-TQVFB-DG9YT-724CC
echo 90c362e5-0da1-4bfd-b53b-b87d309ade43*6NMRW-2C8FM-D24W7-TQWMY-CWH2D
echo a99cc1f0-7719-4306-9645-294102fbff95*FDNH6-VW9RW-BXPJ7-4XTYG-239TB
echo 8de8eb62-bbe0-40ac-ac17-f75595071ea3*GRFBW-QNDC4-6QBHG-CCK3B-2PR88
echo 2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283*JCKRF-N37P4-C2D82-9YXRT-4M63B
echo 8c1c5410-9f39-4805-8c9d-63a07706358f*WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
echo 21c56779-b449-4d20-adfc-eece0e1ad74b*CB7KF-BWN84-R7R2Y-793K2-8XDDG
echo 61c5ef22-f14f-4553-a824-c4b31e84b100*PTXN8-JFHJM-4WC78-MPCBR-9W4KR
echo e49c08e7-da82-42f8-bde2-b570fbcae76c*2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG
echo 7b4433f4-b1e7-4788-895a-c45378d38253*QN4C6-GBJD2-FB422-GHWJK-GJG2R
echo 3dbf341b-5f6c-4fa7-b936-699dce9e263f*VP34G-4NPPG-79JTQ-864T4-R3MQX
echo 43d9af6e-5e86-4be8-a797-d072a046896c*K9FYF-G6NCK-73M32-XMVPY-F9DRR
echo c06b6981-d7fd-4a35-b7b4-054742b7af67*GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
echo 7476d79f-8e48-49b4-ab63-4d0b813a16e4*HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
echo 81671aaf-79d1-4eb1-b004-8cbbe173afea*MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
echo 113e705c-fa49-48a4-beea-7dd879b46b14*TT4HM-HN7YT-62K67-RGRQJ-JFFXW
echo 096ce63d-4fac-48a9-82a9-61ae9e800e5f*789NJ-TQK6T-6XTH8-J39CJ-J8D3P
echo fe1c3238-432a-43a1-8e25-97e7d1ef10f3*M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
echo 78558a64-dc19-43fe-a0d0-8075b2a370a3*7B9N3-D94CG-YTVHR-QBPX3-RJP64
echo ffee456a-cd87-4390-8e07-16146c672fd0*XYTND-K6QKT-K2MRH-66RTM-43JKP
echo c72c6a1d-f252-4e7e-bdd1-3fca342acb35*BB6NG-PQ82V-VRDPW-8XVD2-V8P66
echo db78b74f-ef1c-4892-abfe-1e66b8231df6*NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3
echo 0ab82d54-47f4-4acb-818c-cc5bf0ecb649*NMMPB-38DD4-R2823-62W8D-VXKJB
echo cd4e2d9f-5059-4a50-a92d-05d5bb1267c7*FNFKF-PWTVT-9RC8H-32HB2-JB34X
echo f7e88590-dfc7-4c78-bccb-6f3865b99d1a*VHXM3-NR6FT-RY6RT-CK882-KW2CJ
echo e9942b32-2e55-4197-b0bd-5ff58cba8860*3PY8R-QHNP9-W7XQD-G6DPH-3J2C9
echo c6ddecd6-2354-4c19-909b-306a3058484e*Q6HTR-N24GM-PMJFP-69CD8-2GXKR
echo b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c*KF37N-VDV38-GRRTV-XH8X6-6F3BB
echo ba998212-460a-44db-bfb5-71bf09d1c68b*R962J-37N87-9VVK2-WJ74P-XTMHR
echo e58d87b5-8126-4580-80fb-861b22f79296*MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B
echo cab491c7-a918-4f60-b502-dab75e334f40*TNFGH-2R6PB-8XM3K-QYHX2-J4296
echo b3ca044e-a358-4d68-9883-aaa2941aca99*D2N9P-3P6X9-2R39C-7RTCD-MDVJX
echo 00091344-1ea4-4f37-b789-01750ba6988c*W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
echo 21db6ba4-9a7b-4a14-9e29-64a60c59301d*KNC87-3J2TX-XB4WP-VCPJV-M4FWM
echo b743a2be-68d4-4dd3-af32-92425b7bb623*3NPTF-33KPT-GGBPR-YX76B-39KDD
echo a98bcd6d-5343-4603-8afe-5908e4611112*NG4HW-VH26C-733KW-K6F98-J8CK4
echo ebf245c1-29a8-4daf-9cb1-38dfc608a8c8*XCVCF-2NXM9-723PB-MHCB7-2RYQQ
echo 458e1bec-837a-45f6-b9d5-925ed5d299de*32JNW-9KQ84-P47T8-D8GGY-CWCK7
echo e14997e7-800a-4cf7-ad10-de4b45b578db*JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
echo a00018a3-f20f-4632-bf7c-8daa5351c914*GNBB8-YVD74-QJHX6-27H4K-8QHDG
echo c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60*BN3D2-R7TKB-3YPBD-8DRP2-27GG4
echo 197390a0-65f6-4a95-bdc4-55d58a3b0253*8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
echo 8860fcd4-a77b-4a20-9045-a150ff11d609*2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
echo 9d5584a2-2d85-419a-982c-a00888bb9ddf*4K36P-JN4VD-GDC6V-KDT89-DYFKP
echo af35d7b7-5035-4b63-8972-f0b747b9f4dc*DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV
echo 10018baf-ce21-4060-80bd-47fe74ed4dab*RYXVT-BNQG7-VD29F-DBMRY-HT73M
echo 18db1848-12e0-4167-b9d7-da7fcda507db*NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2
echo f0f5ec41-0d55-4732-af02-440a44a3cf0f*XC9B7-NBPP2-83J2H-RHMBY-92BT4
echo d3643d60-0c42-412d-a7d6-52e6635327f6*48HP8-DN98B-MYWDG-T2DCC-8W83P
echo 7d5486c7-e120-4771-b7f1-7b56c6d3170c*HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
echo 95fd1c83-7df5-494a-be8b-1300e1c9d1cd*XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
echo b92e9980-b9d5-4821-9c94-140f632f6312*FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
echo 54a09a0d-d57b-4c10-8b69-a842d6590ad5*MRPKT-YTG23-K7D7T-X2JMM-QY7MG
echo 5a041529-fef8-4d07-b06f-b59b573b32d2*W82YF-2Q76Y-63HXB-FGJG9-GF7QX
echo ae2ee509-1b34-41c0-acb7-6d4650168915*33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
echo 1cb6d605-11b3-4e14-bb30-da91c8e3983a*YDRBP-3D83W-TY26F-D46B2-XCKRJ
echo 46bbed08-9c7b-48fc-a614-95250573f4ea*C29WB-22CC8-VJ326-GHFJW-H9DH4
echo db537896-376f-48ae-a492-53d0547773d0*YBYF6-BHCR3-JPKRB-CDW7B-F9BK4
echo aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395*73KQT-CD9G6-K7TQG-66MRP-CQ22C
echo e1a8296a-db37-44d1-8cce-7bc961d59c54*XGY72-BRBBT-FF8MH-2GG8H-W7KCW
echo a78b8bd9-8017-4df5-b86a-09f756affa7c*6TPJF-RBVHG-WBW2R-86QPH-6RTM4
echo cda18cf3-c196-46ad-b289-60c072869994*TT8MH-CG224-D3D7Q-498W2-9QCTX
echo 68531fb9-5511-4989-97be-d11a0f55633f*YC6KT-GKW9T-YTKYR-T4X34-R7VHC
echo 620e2b3d-09e7-42fd-802a-17a13652fe7a*489J6-VHDMP-X63PK-3K798-CPX3Y
echo 7482e61b-c589-4b7f-8ecc-46d455ac3b87*74YFP-3QFB3-KQT8W-PMXWJ-7M648
echo 8a26851c-1c7e-48d3-a687-fbca9b9ac16b*GT63C-RJFQ3-4GMB6-BRFB9-CB83V
echo f772515c-0e87-48d5-a676-e6962c3e1195*736RG-XDKJK-V34PF-BHK87-J6X3K
echo 4f3d1606-3fea-4c01-be3c-8d671c401e3b*YFKBB-PQJJV-G996G-VWGXY-2V3X8
echo 2c682dc2-8b68-4f63-a165-ae291d4cf138*HMBQG-8H2RH-C77VX-27R82-VMQBT
echo cfd8ff08-c0d7-452b-9f60-ef5c70c32094*VKK3X-68KWM-X2YGT-QR4M6-4BWMV
echo d4f54950-26f2-4fb4-ba21-ffab16afcade*VTC42-BM838-43QHV-84HX6-XJXKV
echo ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b*WYR28-R7TFJ-3X2YQ-YCY4H-M249D
echo ad2542d4-9154-4c6d-8a44-30f11ee96989*TM24T-X9RMF-VWXK6-X8JC9-BFGM2
echo 2401e3d0-c50a-4b58-87b2-7e794b7d2607*W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
echo c1af4d90-d1bc-44ca-85d4-003ba33db3b9*YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
echo 8198490a-add0-47b2-b3ba-316b12d647b4*39BXF-X8Q23-P2WWT-38T2F-G3FPG
echo 7afb1156-2c1d-40fc-b260-aab7442b62fe*RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
echo 68b6e220-cf09-466b-92d3-45cd964b9509*7M67G-PC374-GR742-YH8V4-TCBY3
echo fd09ef77-5647-4eff-809c-af2b64659a45*22XQ2-VRXRG-P8D42-K34TD-G3QQC
echo 01ef176b-3e0d-422a-b4f8-4ea880035e8f*4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto :eof

:: VBS Helpers
:: VBS Helpers
:: VBS Helpers

:Query

set pStart=%time: =0%

Rem Main Code Here
>nul 2>&1 del /q "%temp%\result"
if /i '%3' EQU '' (
	>"%temp%\result" cscript "%windir%\KmsHelper.vbs" "/QUERY_BASIC" %1 %2
) else (
	>"%temp%\result" cscript "%windir%\KmsHelper.vbs" "/QUERY_ADVENCED" %1 %2 %3
)

Rem Main Code Here

set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
 if /i '%3' EQU '' (
	%printCmd_v2% cscript "%windir%\KmsHelper.vbs" "/QUERY_BASIC" %1 %2
) else (
	%printCmd_v2% cscript "%windir%\KmsHelper.vbs" "/QUERY_ADVENCED" %1 %2 %3
)
%PrintLog% type "%temp%\result"
%PrintLog% echo ******************************************

goto :eof

:ACTIVATE

set pStart=%time: =0%

Rem Main Code Here
set "lastErr="
set "activationCMD=2>nul cscript "%windir%\KmsHelper.vbs" "/ACTIVATE" "%1" "%2""
for /f "tokens=1,2 delims=: " %%x in ('"!activationCMD!"') do set "lastErr=%%y"
Rem Main Code Here

set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
%printCmd_v2% cscript "%windir%\KmsHelper.vbs" "/ACTIVATE" "%1" "%2"
%PrintLog% echo Last Error !lastErr!
%PrintLog% echo ******************************************
goto :eof

:UninstallProductKey

set pStart=%time: =0%

Rem Main Code Here
>nul 2>&1 del /q "%temp%\result"
>"%temp%\result" 2>&1 cscript "%windir%\KmsHelper.vbs" "/UninstallProductKey" %1 %2
Rem Main Code Here

set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
%printCmd_v2% cscript "%windir%\KmsHelper.vbs" "/UninstallProductKey" %1 %2
%PrintLog% type "%temp%\result"
%PrintLog% echo ******************************************
goto :eof

Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION Script
Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION Script
Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION Script

:STARTKMSActivation
set SSppHook=0
set KMSPort=1688
set KMSHostIP=%IP_ADDRESS%
set KMS_RenewalInterval=10080
set KMS_ActivationInterval=120
set KMS_HWID=0x3A1C049600B60076

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"

if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xOS=A64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xOS=x86"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xOS=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xOS=A64"

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set _Hook="%SysPath%\SppExtComObjHook.dll"

for /f %%A in ('"2>nul dir /b /ad %SysPath%\spp\tokens\skus"') do (
	if !buildnumber! GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
	if !buildnumber! LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
	if !buildnumber! LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc >nul 2>&1
if %errorlevel% EQU 1060 set OsppHook=0

set ESU_KMS=0
if !buildnumber! LSS 9200 for /f %%A in ('"2>nul dir /b /ad %SysPath%\spp\tokens\channels"') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
if %ESU_KMS% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "addon=and LicenseDependsOn is not NULL") else (set "adoff="&set "addon=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
if %ESU_EDT% EQU 1 set SSppHook=1
set ESU_ADD=0

if !buildnumber! GEQ 9200 (
	set OSType=Win8
	set SppVer=SppExtComObj.exe
) else if !buildnumber! GEQ 7600 (
	set OSType=Win7
	set SppVer=sppsvc.exe
) else (
	pause
	exit /b
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" >nul 2>&1 && (
	reg delete "%IFEO%\sppsvc.exe" /f >nul 2>&1
	call :StopService sppsvc
)
set _uRI=%KMS_RenewalInterval%
set _uAI=%KMS_ActivationInterval%
set _AUR=0
if exist %_Hook% dir /b /al %_Hook% >nul 2>&1 || (
  reg query "%IFEO%\%SppVer%" /v VerifierFlags >nul 2>&1 && set _AUR=1
  if %SSppHook% EQU 0 reg query "%IFEO%\osppsvc.exe" /v VerifierFlags >nul 2>&1 && set _AUR=1
)

if %_AUR% EQU 0 (
	set KMS_RenewalInterval=43200
	set KMS_ActivationInterval=43200
) else (
	set KMS_RenewalInterval=%_uRI%
	set KMS_ActivationInterval=%_uAI%
)
if !buildnumber! GEQ 9600 (
	reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f >nul 2>&1
	if !buildnumber! EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f >nul 2>&1
)

call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" >nul 2>&1
  if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" >nul 2>&1
)
set AclReset=0
set _cphk=0
if %_AUR% EQU 1 set _cphk=1
if %_cphk% EQU 1 (
	copy /y "%windir%\bin\%xOS%.dll" %_Hook% >nul 2>&1
	goto :skipsym
)
mklink %_Hook% "%windir%\bin\%xOS%.dll" >nul 2>&1
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 goto :E_SYM
icacls %_Hook% /findsid *S-1-5-32-545 2>nul | find /i "SppExtComObjHook.dll" >nul || (
	set AclReset=1
	icacls %_Hook% /grant *S-1-5-32-545:RX >nul 2>&1
)
:skipsym
if %SSppHook% NEQ 0 call :CreateIFEOEntry %SppVer%
if %_AUR% EQU 1 (call :CreateIFEOEntry osppsvc.exe) else (if %OsppHook% NEQ 0 call :CreateIFEOEntry osppsvc.exe)
if %_AUR% EQU 1 if %OSType% EQU Win7 call :CreateIFEOEntry SppExtComObj.exe
if %_AUR% EQU 1 (
	call :UpdateIFEOEntry %SppVer%
	call :UpdateIFEOEntry osppsvc.exe
)
goto :eof

:StopKMSActivation
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %_AUR% EQU 0 call :RemoveHook
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
goto :eof

:StopService
sc query %1 | find /i "STOPPED" >nul || net stop %1 /y >nul 2>&1
sc query %1 | find /i "STOPPED" >nul || sc stop %1 >nul 2>&1
goto :eof


:RemoveHook
if %AclReset% EQU 1 icacls %_Hook% /reset >nul 2>&1
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
	if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" >nul 2>&1
	if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" >nul 2>&1
)
for %%# in (SppExtComObj.exe,sppsvc.exe,osppsvc.exe) do reg query "%IFEO%\%%#" >nul 2>&1 && (
	call :RemoveIFEOEntry %%#
)
if %OSType% EQU Win8 schtasks /query /tn "%_TaskEx%" >nul 2>&1 && (
	schtasks /delete /f /tn "%_TaskEx%" >nul 2>&1
)
goto :eof

:CreateIFEOEntry
reg delete "%IFEO%\%1" /f /v Debugger >nul 2>nul
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" >nul 2>&1
reg add "%IFEO%\%1" /f /v VerifierDebug /t REG_DWORD /d 0x00000000 >nul 2>&1
reg add "%IFEO%\%1" /f /v VerifierFlags /t REG_DWORD /d 0x80000000 >nul 2>&1
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 0x00000100 >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d 1 >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% >nul 2>&1

if /i %1 EQU SppExtComObj.exe if !buildnumber! GEQ 9600 (
	reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" >nul 2>&1
)
goto :eof

:RemoveIFEOEntry
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f >nul 2>nul
goto :eof
)
if %OsppHook% EQU 0 (
reg delete "%IFEO%\%1" /f >nul 2>nul
)
if %OsppHook% NEQ 0 for %%A in (Debugger,VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /v %%A /f >nul 2>nul
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%" >nul 2>&1
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" >nul 2>&1
goto :eof

:UpdateIFEOEntry
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% >nul 2>&1
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% >nul 2>&1
if /i %1 EQU SppExtComObj.exe if !buildnumber! GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" >nul 2>&1
)
if /i %1 EQU sppsvc.exe (
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% >nul 2>&1
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% >nul 2>&1
)

:UpdateOSPPEntry
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMSHostIP%" >nul 2>&1
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMSPort%" >nul 2>&1
)
goto :eof