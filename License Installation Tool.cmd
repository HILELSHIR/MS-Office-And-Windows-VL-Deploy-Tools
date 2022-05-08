@echo off

set tArgs=
set tProduct=
set tYear=
set HideMode=
::::::::::::::::::::::::::::::::::
set tArgs=%1
set tProduct=%2
set tYear=%3
set HideMode=%4
::::::::::::::::::::::::::::::::::
set ScriptMode=WMI
::::::::::::::::::::::::::::::::::
set LocalKms=true
::::::::::::::::::::::::::::::::::
set fileLoc="%~dpfx0"
set tmpFile="%temp%\tmp"
set "SupportedbuildYear=2019, 2021"
::::::::::::::::::::::::::::::::::
set Invalid="Invalid"
set ProductError="ERROR:"
set ProductNotExist="No Instance(s) Available."
set ProductNotFound="Not found"
::::::::::::::::::::::::::::::::::
set logFile="%temp%\log"
set SecLogFile="%temp%\SecLogFile"
set "PrintNul=1>nul 2>nul"
set "PrintLog=1>>%LogFile% 2>>&1"
set "PrintCmd=1>>%LogFile% echo.&1>>%LogFile% echo"
set "PrintCmd_v2=1>>%LogFile% echo"
::::::::::::::::::::::::::::::::::
setlocal enabledelayedexpansion
title License Installation Tool ~ !ScriptMode! ~ Mode

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

rem get rid of the not genuine banner solution by Windows_Addict
	
rem first NAG ~ check for IP address On start up
rem https://forums.mydigitallife.net/threads/kms_vl_all-smart-activation-script.79535/page-180#post-1659178

rem second NAG ~ check if ip address is from range of 0.0.0.0 to ?
rem HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing\LVUXRecords
rem https://forums.mydigitallife.net/threads/kms_vl_all-smart-activation-script.79535/page-237#post-1734148

rem How-to: Generate Random Numbers
rem https://ss64.com/nt/syntax-random.html

Set /a rand_V=(%RANDOM%*20/32768)+1    || Add-in
Set /a rand_A=(%RANDOM%*255/32768)+1   || 192-255
Set /a rand_B=(%RANDOM%*255/32768)+1   || 168-255
Set /a rand_C=(%RANDOM%*255/32768)+1   || 000-255
Set /a rand_D=(%RANDOM%*255/32768)+1   || 000-255

if !rand_A! LSS 192 Set /a rand_A+=192-!rand_A!+!rand_V!
if !rand_B! LSS 168 Set /a rand_B+=168-!rand_B!+!rand_V!
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

if not exist "Tools\KmsHelper.vbs" (
	echo.
	echo Missing Critical Files [KmsHelper.vbs]
	echo.
	pause
	exit /b
)

if defined LocalKms (

	set "x64FileList=A64.dll,x64.dll,x86.dll"
	REM set "x32FileList=vlmcsd.exe,SECOPatcher.dll,FakeClient.exe,libkms32.dll,libkms64.dll,WinDivert.dll,WinDivert.lib,WinDivert32.sys,WinDivert64.sys"
	REM set "x64FileList=vlmcsd.exe,SECOPatcher.dll,FakeClient.exe,libkms64.dll,WinDivert.dll,WinDivert.lib,WinDivert64.sys,A64.dll,x64.dll,x86.dll"
	
	REM for %%# in (!x32FileList!) do if not exist "x32\%%#" set "LocalKms="
	for %%# in (!x64FileList!) do if not exist "Tools\bin\%%#" set "LocalKms="
	
	if not defined LocalKms (
		echo.
		echo Local Activation files Is Missing, Switch back to Online KMS
		timeout 5
		cls
	)
)

set LicensingServiceClass=SoftwareLicensingService
set LicensingProductClass=SoftwareLicensingProduct

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
	title Licence Installation Tool ~ !ScriptMode! ~ Mode
	set LicensingServiceClass=OfficeSoftwareProtectionService
	set LicensingProductClass=OfficeSoftwareProtectionProduct
)

if defined tYear (
	if /i '!tYear!' EQU '2010' (
		set V2010=_2010
		title Licence Installation Tool ~ !ScriptMode! ~ Mode
		set LicensingServiceClass=OfficeSoftwareProtectionService
		set LicensingProductClass=OfficeSoftwareProtectionProduct
	)
)

set xArgs=
set xArgs=%*

if not defined xArgs (
	cls
	echo.
	echo Licence Tool Will Rebuild All Office Licences
	echo And Convert Them Into Volume Licences
	echo.
	echo * Support Multi Versions
	echo * Can Identify / Convert / Activate, Both Msi ^& C2R On same System.
	echo * -Add -Remove Also Upgraded To Work With Multi Versions
	echo.
	echo If you have ~ Activated ~ Mak Or Retail Licence
	echo And You Wish To KEEP IT ##### Use Kms Tool Instead #####
	echo.
	echo Optional Parameters
	echo "%~n0" -Add Word 2021
	echo "%~n0" -Remove Word 2021
	echo "%~n0" -Convert Mondo 2016
	echo "%~n0" -Convert Standard 2019
	echo "%~n0" -Convert ProPlus 2019
	echo.
	CHOICE /C CE /M "Press C For Continue, E for Exit."
	if !errorlevel! EQU 2 exit /b
	cls
)

for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (
	call :CmdWorker reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Microsoft\Office" /f
	call :CmdWorker reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Wow6432Node\Microsoft\Office" /f
	call :CmdWorker reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides" /f
)

call :GetInfo
call :CheckSystemBit

if not defined xArgs (
	cls
	call :GetProductYear
	goto :end
)

if defined xArgs (
	call :AdditionFunction
	goto :end
)

:end
if not defined HideMode echo.

call :CleanRegistryKeys
call :CmdWorker reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
call :CmdWorker reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
call :CmdWorker reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"

setlocal disabledelayedexpansion
if not defined HideMode echo Log file saved at :: "%temp%\log"
if not defined HideMode echo please wait 6 seconds..
if not defined HideMode timeout /t 6 /NOBREAK
goto :eof

:cleanVariable
set vars=key, c2r_14_Key, c2r_15_Key, root_14, root_15, root_16, guid_14, guid_15, guid_16, cscript, ProductYear, KmsServer, V2010, UnInstallKeys, U_2010, U_2013, U_2016, U_2019, U_2021
for %%v in (%vars%) do (set %%v=)
goto :eof

:CmdWorker
set pStart=%time: =0%
%PrintNul% del /q %SecLogFile%
1>>%SecLogFile% 2>>&1 %*
echo %* |>nul find /i "/" && set lastErrorlvl=%ErrorLevel%
set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
%printCmd_v2% %*
if exist %SecLogFile% %PrintLog% type %SecLogFile%
%PrintLog% echo ******************************************

goto :eof

:CheckSystemBit
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set xBit=32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set xBit=64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set xBit=64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set xBit=64

if exist "%windir%\system32\cscript.exe" set cscript="%windir%\system32\cscript.exe"
if exist "%windir%\SysWOW64\cscript.exe" set cscript="%windir%\SysWOW64\cscript.exe"
goto :eof

:LicenceWorker
if not defined ProductYear (
	if not defined tYear (
		if not defined xArgs (
			echo Could Not find any Supported Office Products
			goto :eof
		)
	)
)

if not defined OfficeMsi_14 (
	if not defined OfficeMsi_15 (
		if not defined OfficeMsi_16 (
		
			rem Must Check if we have something
			rem Must Check if we have something
			rem Must Check if we have something
			
			if defined Officec2r_v15 (
				set "ProductReleaseIds="
				%printCmd% reg query %c2r_15_Key% /v ProductReleaseIds
				%PrintLog% reg query %c2r_15_Key% /v ProductReleaseIds && (
					%PrintLog% echo ******************************************
					for /f "tokens=2 delims=_" %%g in ('reg query %c2r_15_Key% /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set "ProductReleaseIds=%%g"
					set "ProductReleaseIds=!ProductReleaseIds:~6!"
				) || (%PrintLog% echo ******************************************)
			)
		
			if defined Officec2r_v16 (
				set "ProductReleaseIds="
				%printCmd% reg query %c2r_16_Key% /v ProductReleaseIds
				%PrintLog% reg query %c2r_16_Key% /v ProductReleaseIds && (
					%PrintLog% echo ******************************************
					for /f "tokens=2 delims=_" %%g in ('reg query %c2r_16_Key% /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set "ProductReleaseIds=%%g"
					set "ProductReleaseIds=!ProductReleaseIds:~6!"
				) || (%PrintLog% echo ******************************************)
			)
			
			if not defined ProductReleaseIds (
				goto :eof
			)
		)
	)
)
		
set ProductsList=365, HOME, Professional, Private
for %%x in (!ProductsList!) do (

	set SelectedX=
	if defined OfficeMsi_14 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_15 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_16 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r_v15 (
		echo !ProductReleaseIds! | >nul find /i "%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r_v16 (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
		)
	)
	if defined SelectedX (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
			for %%m in (!ProductReleaseIds!) do (
				echo %%m | >nul find /i "%%x" && (
					set ProductYear=
					echo %%m | >nul find /i "2021" && set ProductYear=2021
					echo %%m | >nul find /i "2019" && set ProductYear=2019
					if not defined ProductYear 		  set ProductYear=2016
				)
			)
		)
		
		if /i '!ProductYear!' EQU '2010' (
			call :Convert ProPlus !ProductYear!
			goto :eof
		)
		
		if /i '!ProductYear!' EQU '2013' (
			call :Convert mondo !ProductYear!
			goto :eof
		)
		
		call :Convert mondo 2016
		goto :eof
	)
)
		
set ProductsList=SmallBusBasics,proplus,Standard,mondo,word,excel,powerpoint,Skype,access,outlook,publisher,ProjectPro,ProjectStd,VisioStd,VisioPro,VisioPrem,OneNote
for %%x in (!ProductsList!) do (
	
	set SelectedX=
	if defined OfficeMsi_14 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_15 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_16 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r_v15 (
		echo !ProductReleaseIds! | >nul find /i "%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r_v16 (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
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
	if defined SelectedX (
		if not defined HideMode echo.
		echo Install Volume Licence serials for %%x !ProductYear!
		echo .................................................
		
		call :integrate %%x !ProductYear!
		for /f %%y in ('dir /b licence\%%x*!ProductYear!*.xrm-ms') do call :inslic %%y %%x
		
		if defined Officec2r_v16 (
			call :UnInstallKeys 2016
			call :UnInstallKeys 2019
			call :UnInstallKeys 2021
		) else (
			call :UnInstallKeys !ProductYear!
		)
		
		if defined OfficeMsi_14 (
			if defined Office14_WOW (
				echo Import [Wow] Keys
				call :CmdWorker reg import Keys\%%x\RegistrationWOW.reg
			) else (
				if /i '!xBit!' EQU '32' (
					echo Import [X32] Keys
					call :CmdWorker reg import Keys\%%x\Registration32.reg
				)
				if /i '!xBit!' EQU '64' (
					echo Import [X32] Keys
					echo Import [X64] Keys
					call :CmdWorker reg import Keys\%%x\Registration32.reg
					call :CmdWorker reg import Keys\%%x\Registration64.reg
				)
			)
		)
		
		:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
		>%tmpFile% call :volume_license_serial_list
		for /f "tokens=1,2,3 delims=*" %%p in ('type %tmpFile%') do (
			if /i '%%q' EQU '!ProductYear!' (
				echo %%p | find /i "%%x">nul && call :inpkey %%r
			)
		)
	)
)

goto :eof

:AdditionFunction
if /i '%tArgs%' EQU '-Convert' (
	if /i '!tProduct!' EQU '' (
		cls
		echo product not exist .. wtf 
		goto :eof
	)

	if /i '!tYear!' EQU '' (
		cls
		echo product not exist .. wtf 
		goto :eof
	)
	
	if /i '!HideMode!' EQU '-hide' (title Working ...&set HideMode=true) else (cls&set HideMode=)

	set foundYear=
	for %%g in (2010, 2013, 2016, 2019, 2021) do (if /i '!tYear!' EQU '%%g' (set foundYear=true))
	if not defined foundYear (
		echo product not exist .. wtf
		goto :eof
	)
	
	%printCmd% dir /b licence\!tProduct!*!tYear!*.xrm-ms
	(%PrintLog% dir /b licence\!tProduct!*!tYear!*.xrm-ms) || (
		%PrintLog% echo ******************************************
		echo product not exist .. wtf 
		goto :eof
	)
	%PrintLog% echo ******************************************
	call :GetProductYear
	goto :eof
)

if /i '%tArgs%' EQU '-Add' (

	cls
	if not defined HideMode echo.

	if /i '!tProduct!' EQU '' (
		echo product not exist .. wtf 
		goto :eof
	)
	
	if /i '!tYear!' EQU '' (
		echo product not exist .. wtf 
		goto :eof
	)
	
	set foundYear=
	for %%g in (2010, 2013, 2016, 2019, 2021) do (if /i '!tYear!' EQU '%%g' (set foundYear=true))
	if not defined foundYear (
		echo product not exist .. wtf
		goto :eof
	)
	
	%printCmd% dir /b licence\!tProduct!*!tYear!*.xrm-ms
	(%PrintLog% dir /b licence\!tProduct!*!tYear!*.xrm-ms) || (
		%PrintLog% echo ******************************************
		echo product not exist .. wtf 
		goto :eof
	)
	%PrintLog% echo ******************************************
	
	echo Add Office !tProduct! !tYear! Volume Licence
	echo ................................................
	set "ProductYear=!tYear!"
	call :integrate !tProduct! !ProductYear!
	for /f %%y in ('dir /b licence\!tProduct!*!tYear!*.xrm-ms') do call :inslic %%y !tProduct!
	
	if /i '!ProductYear!' EQU '2010' (
	
		set Office14_WOW=
		%printCmd% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration"
		%PrintLog% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration" && set Office14_WOW=true

		if defined Office14_WOW (
			echo Import [Wow] Keys
			call :CmdWorker reg import Keys\!tProduct!\RegistrationWOW.reg
		) else (
			if /i '!xBit!' EQU '32' (
				echo Import [X32] Keys
				call :CmdWorker reg import Keys\!tProduct!\Registration32.reg
			)
			if /i '!xBit!' EQU '64' (
				echo Import [X32] Keys
				echo Import [X64] Keys
				call :CmdWorker reg import Keys\!tProduct!\Registration32.reg
				call :CmdWorker reg import Keys\!tProduct!\Registration64.reg
			)
		)
	)
	
	:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
	>%tmpFile% call :volume_license_serial_list
	for /f "tokens=1,2,3 delims=*" %%p in ('type %tmpFile%') do (
		if /i '%%q' EQU '!tYear!' (
			echo %%p | find /i "!tProduct!">nul && call :inpkey %%r
		)
	)
	
	if not defined HideMode echo.
	call :Activate !ScriptMode! !ProductYear!
	goto :eof
)

if /i '%tArgs%' EQU '-Remove' (
	
	cls
	if not defined HideMode echo.
	
	if /i '!tProduct!' EQU '' (
		echo product not exist .. wtf 
		goto :eof
	)
	
	if /i '!tYear!' EQU '' (
		echo product not exist .. wtf 
		goto :eof
	)
	
	set foundYear=
	for %%g in (2010, 2013, 2016, 2019, 2021) do (if /i '!tYear!' EQU '%%g' (set foundYear=true))
	
	if not defined foundYear (
		echo product not exist .. wtf
		goto :eof
	)

	echo Remove Office !tProduct! !tYear! Volume Licence
	echo ................................................
	
	set ProductYear=!tYear!
	call :UnInstallKey !tProduct! !tYear!
	goto :eof
)

cls
echo.
echo Bad Args
goto :eof

:GetProductYear

rem Clear Variables
Set OfficeMsi_14=
Set OfficeMsi_15=
Set OfficeMsi_16=
Set Officec2r_v14=
Set Officec2r_v15=
set Officec2r_v16=
set ActivateStatus=

rem C2R 16.X Check

if defined c2r_16_Key (

	if not defined Windows_7_Or_Earlier (
		set LicensingServiceClass=SoftwareLicensingService
		set LicensingProductClass=SoftwareLicensingProduct
	)
	
	set V2010=
	set ProductReleaseIds=
	for /f "tokens=2 delims=_" %%g in ('reg query !c2r_16_Key! /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
	set ProductReleaseIds=!ProductReleaseIds:~6!
	
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office C2R [2016, 2019, 2021] Products
	if not defined HideMode echo --- winner winner chicken dinner
	set Officec2r_v16=true
	call :update_Year xxx
	
	if not defined xArgs (
		call :LicenceWorker
	)
	if defined xArgs (
		set ProductMatch=
		set ProductList=2016, 2019, 2021
		for %%k in (!ProductList!) do (
			if /i '!tYear!' EQU '%%k' set ProductMatch=true
		)
		if defined ProductMatch (
			call :convert !tProduct! !tYear!
		)
		if not defined ProductMatch (
			echo --- Product Not Supported :: !tProduct! !tYear!
		)
	)
	
	set Officec2r_v16=
	if not defined xArgs set ActivateStatus=true
	
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office C2R [2016, 2019, 2021] Products
	if not defined HideMode echo --- 404 not found
)

rem C2R 15.X Check

if defined c2r_15_Key (

	if not defined Windows_7_Or_Earlier (
		set LicensingServiceClass=SoftwareLicensingService
		set LicensingProductClass=SoftwareLicensingProduct
	)
	
	set V2010=
	set ProductReleaseIds=
	for /f "tokens=2 delims=_" %%g in ('reg query !c2r_15_Key! /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
	set ProductReleaseIds=!ProductReleaseIds:~6!
	
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office C2R [2013] Products
	if not defined HideMode echo --- winner winner chicken dinner
	set Officec2r_v15=true
	call :update_Year 2013
	
	if not defined xArgs (
		call :LicenceWorker
	)
	if defined xArgs (
		set ProductMatch=
		set ProductList=2013
		for %%k in (!ProductList!) do (
			if /i '!tYear!' EQU '%%k' set ProductMatch=true
		)
		if defined ProductMatch (
			call :convert !tProduct! !tYear!
		)
		if not defined ProductMatch (
			echo --- Product Not Supported :: !tProduct! !tYear!
		)
	)
	
	set Officec2r_v15=
	if not defined xArgs set ActivateStatus=true
	
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office C2R [2013] Products
	if not defined HideMode echo --- 404 not found
)

rem C2R 14.X Check

if defined c2r_14_Key (
	echo.
	call :update_Year 2010
	echo ##### Search For Office C2R [2010] Products
	echo --- winner winner chicken dinner
	
	rem Actully there is Some Products, But i dont have licence OR Serial for them
	rem Office Home and Business 2010 \ Office Home and Student 2010 \ Office Starter 2010
	rem https://docs.microsoft.com/en-us/officeupdates/update-history-office-2010-click-to-run

	echo.
	echo Please Upgrade to 2013 and above
	echo If you Insist keep using Office 2010
	echo Install [Msi] Version Instead, This one Supported
	explorer "https://the-eye.eu/public/MSDN/Office%202010/"
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office C2R [2010] Products
	if not defined HideMode echo --- 404 not found
)

rem MSI 16.X Check

(reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16") && Set OfficeMsi_16=true
(reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16") && Set OfficeMsi_16=true

if defined OfficeMsi_16 (
	if not defined Windows_7_Or_Earlier (
		set LicensingServiceClass=SoftwareLicensingService
		set LicensingProductClass=SoftwareLicensingProduct
	)
	echo.
	set V2010=
	echo ##### Search For Office MSI [2016] Products
	echo --- winner winner chicken dinner
	call :update_Year 2016
	
	if not defined xArgs (
		call :LicenceWorker
	)
	if defined xArgs (
		set ProductMatch=
		set ProductList=2016
		for %%k in (!ProductList!) do (
			if /i '!tYear!' EQU '%%k' set ProductMatch=true
		)
		if defined ProductMatch (
			set pKey=
			set SPPSkuId=
			set OfficeVer=office16
			reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			for /f "tokens=*" %%g in ('reg query !pKey! ^|find /i "!OfficeVer!"') do (
				for /f "tokens=3 delims= " %%h in ('reg query "%%g" /v SPPSkuId') do (
					if not defined SPPSkuId set SPPSkuId=%%h
				)
			)
			call :convert !tProduct! !tYear!
		)
		if not defined ProductMatch (
			echo --- Product Not Supported :: !tProduct! !tYear!
		)
	)
	
	set OfficeMsi_16=
	if not defined xArgs set ActivateStatus=true
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office MSI [2016] Products
	if not defined HideMode echo --- 404 not found
)

rem MSI 15.X Check

(reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15") && Set OfficeMsi_15=true
(reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15") && Set OfficeMsi_15=true

if defined OfficeMsi_15 (
	if not defined Windows_7_Or_Earlier (
		set LicensingServiceClass=SoftwareLicensingService
		set LicensingProductClass=SoftwareLicensingProduct
	)
	echo.
	set V2010=
	echo ##### Search For Office MSI [2013] Products
	echo --- winner winner chicken dinner
	call :update_Year 2013
	
	if not defined xArgs (
		call :LicenceWorker
	)
	if defined xArgs (
		set ProductMatch=
		set ProductList=2013
		for %%k in (!ProductList!) do (
			if /i '!tYear!' EQU '%%k' set ProductMatch=true
		)
		if defined ProductMatch (
			set pKey=
			set SPPSkuId=
			set OfficeVer=office15
			reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			for /f "tokens=*" %%g in ('reg query !pKey! ^|find /i "!OfficeVer!"') do (
				for /f "tokens=3 delims= " %%h in ('reg query "%%g" /v SPPSkuId') do (
					if not defined SPPSkuId set SPPSkuId=%%h
				)
			)
			call :convert !tProduct! !tYear!
		)
		if not defined ProductMatch (
			echo --- Product Not Supported :: !tProduct! !tYear!
		)
	)
	
	set OfficeMsi_15=
	if not defined xArgs set ActivateStatus=true
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office MSI [2013] Products
	if not defined HideMode echo --- 404 not found
)

if defined ActivateStatus (
	if not defined HideMode echo.
	call :Activate !ScriptMode!
	set ActivateStatus=
)

rem MSI 14.X Check

(reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "office14") && Set OfficeMsi_14=true
(reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "office14") && Set OfficeMsi_14=true

if defined OfficeMsi_14 (
	echo.
	echo ##### Search For Office MSI [2010] Products
	echo --- winner winner chicken dinner
	set V2010=_2010
	set LicensingServiceClass=OfficeSoftwareProtectionService
	set LicensingProductClass=OfficeSoftwareProtectionProduct
	title Licence Installation Tool ~ !ScriptMode! ~ Mode
	call :update_Year 2010
	
	if not defined xArgs (
		call :LicenceWorker
	)
	if defined xArgs (
		set ProductMatch=
		set ProductList=2010
		for %%k in (!ProductList!) do (
			if /i '!tYear!' EQU '%%k' set ProductMatch=true
		)
		if defined ProductMatch (
			set pKey=
			set SPPSkuId=
			set OfficeVer=office14
			reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
			reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "!OfficeVer!" && set pKey="HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
			for /f "tokens=*" %%g in ('reg query !pKey! ^|find /i "!OfficeVer!"') do (
				for /f "tokens=3 delims= " %%h in ('reg query "%%g" /v SPPSkuId') do (
					if not defined SPPSkuId set SPPSkuId=%%h
				)
			)
			call :convert !tProduct! !tYear!
		)
		if not defined ProductMatch (
			echo --- Product Not Supported :: !tProduct! !tYear!
		)
	)
	
	set OfficeMsi_14=
	if not defined xArgs set ActivateStatus=true
) else (
	if not defined HideMode echo.
	if not defined HideMode echo ##### Search For Office MSI [2010] Products
	if not defined HideMode echo --- 404 not found
)

if defined ActivateStatus (
	echo.
	call :Activate !ScriptMode!
)

goto :eof

:BuildInfoPharser
if /i '%1' EQU 'ProductYear' ( 
	for %%g in (%SupportedbuildYear%) do (
		if /i '%%g' EQU '%2' call :update_Year %%g
	)
)
goto :eof

:update_Year
set ProductYear=%*
goto :eof

:integrate

set "root="
set "guid="
set "nYear=%2"
set "nProduct=%1"

echo '%nYear%' | find /i "2010">nul && set nProduct=%nProduct%Volume.14
echo '%nYear%' | find /i "2013">nul && set nProduct=%nProduct%Volume.15
echo '%nYear%' | find /i "2016">nul && set nProduct=%nProduct%Volume.16
echo '%nYear%' | find /i "2019">nul && set nProduct=!nProduct!!nYear!Volume.16
echo '%nYear%' | find /i "2021">nul && set nProduct=!nProduct!!nYear!Volume.16

echo '%nYear%' | find /i "2010">nul && (
	if defined OfficeMsi_14 (if not defined SPPSkuId goto :eof)
	if defined Officec2r_v14 (if not defined guid_14 goto :eof)
	
	rem it seems msi version not have integrator
	if not defined root_14 goto :eof
	
	if defined Officec2r_v14 set "guid=%guid_14%"
	if defined OfficeMsi_14 set "guid=%SPPSkuId%"
	set "root=%root_14%"
)

echo '%nYear%' | find /i "2013">nul && (
	if defined OfficeMsi_15 (if not defined SPPSkuId goto :eof)
	if defined Officec2r_v15 (if not defined guid_15 goto :eof)
	
	rem it seems msi version not have integrator
	if not defined root_15 goto :eof
	
	if defined Officec2r_v15 set "guid=%guid_15%"
	if defined OfficeMsi_15 set "guid=%SPPSkuId%"
	set "root=%root_15%"
)

if not defined guid (
	if defined OfficeMsi_16 (if not defined SPPSkuId goto :eof)
	if defined Officec2r_v16 (if not defined guid_16 goto :eof)
	
	rem it seems msi version not have integrator
	if not defined root_16 goto :eof
	
	if defined Officec2r_v16 set "guid=%guid_16%"
	if defined OfficeMsi_16 set "guid=%SPPSkuId%"
	set "root=%root_16%"
)

set SPPSkuId=
if not defined HideMode echo Integrate !nProduct! licence
call :CmdWorker "%root%\Integration\integrator" /I /License PRIDName=!nProduct! PackageGUID="!guid!" /Global /C2R PackageRoot="!root!"
goto :eof

:inslic
if not defined HideMode echo install license file for %2 !ProductYear! :: %1
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /inslic:"Licence\%1"
goto :eof

:inpkey
set xKey=%1
if not defined HideMode echo Install Key :: %xKey:~-5%
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /inpkey:%xKey%
goto :eof

:unpkey
set xKey=%1
if not defined HideMode echo Uninstall Key :: %xKey%
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /unpkey:%xKey%
goto :eof

:UnInstallKeys

:: function to remove
:: all office keys

set tArgs=%*

echo %tArgs% | >nul find /i "2010" && (
	if defined U_2010 goto :eof
	set U_2010=Done
)

echo %tArgs% | >nul find /i "2013" && (
	if defined U_2013 goto :eof
	set U_2013=Done
)

echo %tArgs% | >nul find /i "2016" && (
	if defined U_2016 goto :eof
	set U_2016=Done
)

echo %tArgs% | >nul find /i "2019" && (
	if defined U_2019 goto :eof
	set U_2019=Done
)

echo %tArgs% | >nul find /i "2021" && (
	if defined U_2021 goto :eof
	set U_2021=Done
)

if defined U_2010 (
	set Office14_WOW=
	%printCmd% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration"
	%PrintLog% reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration" && set Office14_WOW=true
	%PrintLog% echo ******************************************
	call :CmdWorker reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Registration" /f
	call :CmdWorker reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Registration" /f
)

if defined tArgs (
	
	if /i '%tArgs%' EQU '2010' (
		call :UnInstallKeys_Worker OfficeSoftwareProtectionProduct %tArgs%
		goto :eof
	)
	
	if defined Windows_7_Or_Earlier (
		call :UnInstallKeys_Worker OfficeSoftwareProtectionProduct %tArgs%
		goto :eof
	)
	
	call :UnInstallKeys_Worker SoftwareLicensingProduct %tArgs%
	goto :eof
)

for %%# in (SoftwareLicensingProduct, OfficeSoftwareProtectionProduct) do call :UnInstallKeys_Worker %%#
goto :eof

:UnInstallKeys_Worker
set xYear=
set xClass=

set xClass=%1
set xYear=%2

call :CmdWorker del /q %tmpFile%

if defined xYear (
	rem Show office 2010 AS Office 14
	echo !xYear! |>nul find /i "2010" && set xYear=2014
	
	rem Show office 2013 AS Office 15
	echo !xYear! |>nul find /i "2013" && set xYear=2015
)

if defined xYear (
	call :query "PartialProductKey, ID" "%xClass%" "Name like '%%%%office !xYear:~-2!%%%%' and PartialProductKey is not null"
	>>%tmpFile% type "%temp%\result"
)
if not defined xYear (
	call :query "PartialProductKey, ID" "%xClass%" "Name like '%%%%Office%%%%' and PartialProductKey is not null"
	>>%tmpFile% type "%temp%\result"
)

set Product_Not_Found=
(type %tmpFile%|find /i ",">nul) 				|| set Product_Not_Found=true
(type %tmpFile%|find /i %invalid%>nul) 			&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductNotExist%>nul) 	&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductNotFound%>nul) 	&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductError%>nul) 	&& set Product_Not_Found=true

if defined Product_Not_Found goto :eof

for /f "tokens=1,2 skip=3 delims=," %%x in ('type %tmpFile%') do (
	:: search for un needed items to avoid echo. that break my for loop
	if /i '%%x' NEQ '' (
		if not defined HideMode echo Remove Key  :: %%y
	
		if /i "!xYear!" EQU "2014" (
			call :CmdWorker %cscript% Tools\x!xBit!\ospp_2010.vbs /unpkey:%%y
			type %SecLogFile% | >nul find /i "not found" && echo ERROR ##### Fail to remove key
		) else (
			call :CmdWorker %cscript% Tools\x!xBit!\ospp.vbs /unpkey:%%y
			type %SecLogFile% | >nul find /i "not found" && echo ERROR ##### Fail to remove key
		)
	)
)
goto :eof

:UnInstallKey

:: function to remove
:: specific office KMS VL keys

set xProduct=%1
set xYear=%2
set xClass=

if /i '!xYear!' EQU '2010' 			set xClass=OfficeSoftwareProtectionProduct
if defined Windows_7_Or_Earlier 	set xClass=OfficeSoftwareProtectionProduct
if not defined xClass				set xClass=SoftwareLicensingProduct

rem Show office 2010 AS Office 14
echo !xYear! |>nul find /i "2010" && set xYear=2014

rem Show office 2013 AS Office 15
echo !xYear! |>nul find /i "2013" && set xYear=2015

call :CmdWorker del /q %tmpFile%
set wmiSearch=PartialProductKey is not null and Name like '%%%%office !xYear:~2!%%%%' And Name like '%%%%!xProduct!%%%%'
call :query "PartialProductKey, ID" "!xClass!" "!wmiSearch!"
>>%tmpFile% type "%temp%\result"

set Product_Not_Found=
(type %tmpFile%|find /i ",">nul) 				|| set Product_Not_Found=true
(type %tmpFile%|find /i %invalid%>nul) 			&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductNotExist%>nul) 	&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductNotFound%>nul) 	&& set Product_Not_Found=true
(type %tmpFile%|find /i %ProductError%>nul) 	&& set Product_Not_Found=true

if defined Product_Not_Found goto :eof

for /f "tokens=1,2 skip=3 delims=," %%x in ('type %tmpFile%') do (
	:: search for un needed items to avoid echo. that break my for loop
	if /i '%%x' NEQ '' (
		if not defined HideMode echo Remove Key  :: %%y
		
		if /i "!xYear!" EQU "2014" (
			call :CmdWorker %cscript% Tools\x!xBit!\ospp_2010.vbs /unpkey:%%y
			type %SecLogFile% | >nul find /i "not found" && echo ERROR ##### Fail to remove key
		) else (
			call :CmdWorker %cscript% Tools\x!xBit!\ospp.vbs /unpkey:%%y
			type %SecLogFile% | >nul find /i "not found" && echo ERROR ##### Fail to remove key
		)
	)
)

goto :eof

:Activate

if defined LocalKms (

	if not defined HideMode echo Local activation
	if not defined HideMode echo .................
	
	call :CleanRegistryKeys
	call :StartKMSActivation
	
	call :UpdateRegistryKeys !KMSHostIP! !KMSPort!
	call :Activate_%1 !KMSHostIP! !KMSPort! %2

	call :CleanRegistryKeys
	call :StopKMSActivation
	goto :eof
	
) else (

	if not defined HideMode echo Online activation
	if not defined HideMode echo .................
	>%tmpFile% call :Kms_Servers_List
	for /f "tokens=*" %%g in ('type %tmpFile%') do (
		if not defined HideMode echo Check if %%g:1688 is Online
		%printCmd% tools\tcping -4 -n 1 -g 1 -w 0.5 -i 0.5 %%g 1688
		%PrintLog% tools\tcping -4 -n 1 -g 1 -w 0.5 -i 0.5 %%g 1688 && (
			%PrintLog% echo ******************************************
			if not defined HideMode echo Winner Winner Chicken dinner
			
			call :Activate_%1 %%g 1688 %2
			goto :eof
		)
		%PrintLog% echo ******************************************
	)

	echo.
	echo didnt found any online kms server
	goto :eof
)

goto :eof

:Activate_VBS
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /sethst:%1
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /setprt:%2
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /act
if not defined HideMode  (
	(type %SecLogFile% | find /i "Product activation successful">nul) && (echo activation succeeded) || (echo activation failed)
)
call :CmdWorker %cscript% Tools\x!xBit!\ospp!V2010!.vbs /remhst
goto :eof

:Activate_WMI

rem Office 2010 Old ID 				:: 59a52881-a989-479d-af46-f275c6370663
if /i '%ProductYear%' EQU '2010' (
	call :Search_14.X_VL_Products
	goto :Activate_WMI_Product_List
)

rem Office 2013 and above New id 	:: 0ff1ce15-a989-479d-af46-f275c6370663
call :Search_VL_Products %3

:Activate_WMI_Product_List

set Product_Not_Found=
(type "%temp%\result"|find /i ",">nul) 					|| set Product_Not_Found=true
(type "%temp%\result"|find /i %invalid%>nul) 			&& set Product_Not_Found=true
(type "%temp%\result"|find /i %ProductError%>nul)		&& set Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotExist%>nul) 	&& set Product_Not_Found=true
(type "%temp%\result"|find /i %ProductNotFound%>nul) 	&& set Product_Not_Found=true

if defined Product_Not_Found (goto :eof)

for /f "tokens=1,2,3,4,5,6,7,8 skip=3 delims=," %%g in ('type "%temp%\result"') do (
	
	echo.
	set "editionName=%%l"
	set /a "Period=%%h/60/24"
	
	echo !editionName:~1, -8!
	echo ....................................
	echo Partial Product Key	= %%m
	echo License / Genuine 	= %%j / %%g
	echo Period Remaining 	= !Period! Days
	echo Id 			= %%i
	echo.
	echo ^+^+^+ Activating ^+^+^+ && echo ...................
	call :Activate_WMI_ %1 %2 %%i %%h
)

goto :eof

:Activate_WMI_

set id=%3
set SPP_ACT_CLASS=
set SPP_KMS_Class=
set SPP_KMS_Where=

set subKey=0ff1ce15-a989-479d-af46-f275c6370663

if defined Windows_7_Or_Earlier (

	rem Office in windows Windows 7 and less use 2 classes
	rem OfficeSoftwareProtectionService for KMS settings 
	rem OfficeSoftwareProtectionProduct for activation
	
	set SPP_KMS_Class=OfficeSoftwareProtectionService
	set SPP_KMS_Where=version is not null
	set SPP_ACT_CLASS=OfficeSoftwareProtectionProduct
)

if /i '%productYear%' EQU '2010' (

	rem Office 2010 Classes
	rem OfficeSoftwareProtectionService for KMS settings 
	rem OfficeSoftwareProtectionProduct for activation
	
	set subKey=59a52881-a989-479d-af46-f275c6370663
	set SPP_KMS_Class=OfficeSoftwareProtectionService
	set SPP_KMS_Where=version is not null
	set SPP_ACT_CLASS=OfficeSoftwareProtectionProduct
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

call :UpdateRegistryKeys %1 %2

:: activation always using SoftwareLicensingProduct
call :ACTIVATE_VBX !Product_Licensing_Class! !Id!

REM :: compare old time to new time
call :Query "GracePeriodRemaining" "!Product_Licensing_Class!" "!Product_Licensing_Where!"
for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set GracePeriod=%%g
set GracePeriod=!GracePeriod: =!

echo Old Grace               = %4
echo New Grace               = !GracePeriod!
if not defined lastErr echo Status                  = Unknown
if defined lastErr if /i '!lastErr!' NEQ '0' (echo Status                  = Failed ^(Error 0x%lastErr%^)) else (echo Status                  = Succeeded ^(Error %lastErr%^))
echo !SPP_KMS_Class! 			| >nul find /i "Office" 	&& (echo Product Class           = OfficeSoftwareProtectionService) || (echo Product Class           = SoftwareLicensingProduct)
echo !Product_Licensing_Class! 	| >nul find /i "Office" 	&& (echo Licensing Class         = OfficeSoftwareProtectionProduct) || (echo Licensing Class         = SoftwareLicensingProduct)
goto :eof

:Search_VL_Products

:: main function to find any
:: VL product with Serial
:: and with custom search support
:: can be Office or Windows

set xYear=
if /i '%*' NEQ '' (
	set xYear=%*
	if /i !xYear! EQU 2013 set xYear=2015
)

call :CmdWorker del /q %tmpFile%
set info=ID, LicenseStatus, PartialProductKey, GenuineStatus, Name, GracePeriodRemaining
set wmiSearch=ApplicationId like '%%%%0ff1ce15-a989-479d-af46-f275c6370663%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null
if defined xYear set wmiSearch=Name like '%%%%office !xYear:~-2!%%%%' and ApplicationId like '%%%%0ff1ce15-a989-479d-af46-f275c6370663%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null
call :Query "!info!" "!LicensingProductClass!" "!wmiSearch!"
>>%tmpFile% type "%temp%\result"
goto :eof

:Search_14.X_VL_Products
call :CmdWorker del /q %tmpFile%
set ApplicationId=59a52881-a989-479d-af46-f275c6370663
set LicensingProductClass=OfficeSoftwareProtectionProduct
set info=ID, LicenseStatus, PartialProductKey, GenuineStatus, Name, GracePeriodRemaining
set wmiSearch=ApplicationId like '%%%%!ApplicationId!%%%%' and Description like '%%%%KMS%%%%' and PartialProductKey is not null
call :Query "!info!" "!LicensingProductClass!" "!wmiSearch!"
>>%tmpFile% type "%temp%\result"
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

:GetInfo

set ROOT_14_X64="%ProgramFiles%\Microsoft Office 14\root"
set ROOT_14_X32="%ProgramFiles(x86)%\Microsoft Office 14\root"
set ROOT_15_X64="%ProgramFiles%\Microsoft Office 15\root"
set ROOT_15_X32="%ProgramFiles(x86)%\Microsoft Office 15\root"
set ROOT_16_X64="%ProgramFiles%\Microsoft Office\root"
set ROOT_16_X32="%ProgramFiles(x86)%\Microsoft Office\root"

if exist %ROOT_14_X64% set "root_14=%ROOT_14_X64:~1,-1%"
if exist %ROOT_14_X32% set "root_14=%ROOT_14_X32:~1,-1%"
if exist %ROOT_15_X64% set "root_15=%ROOT_15_X64:~1,-1%"
if exist %ROOT_15_X32% set "root_15=%ROOT_15_X32:~1,-1%"
if exist %ROOT_16_X64% set "root_16=%ROOT_16_X64:~1,-1%"
if exist %ROOT_16_X32% set "root_16=%ROOT_16_X32:~1,-1%"

set C2R_14_X64_SK="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\ClickToRun\propertyBag"
set C2R_14_X32_SK="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\ClickToRun\propertyBag"
set C2R_15_X64_SK="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
set C2R_15_X32_SK="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag"
set C2R_16_X64_SK="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag"
set C2R_16_X32_SK="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag"

%printCmd% reg query %C2R_14_X64_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X64_SK% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X64_SK%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_14_X32_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X32_SK% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X32_SK%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X64_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X64_SK% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X64_SK%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X32_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X32_SK% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X32_SK%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X64_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X64_SK% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X64_SK%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X32_SK% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X32_SK% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X32_SK%
%PrintLog% echo ******************************************

set C2R_14_X64_ST="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\ClickToRunStore"
set C2R_14_X32_ST="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\ClickToRunStore"
set C2R_15_X64_ST="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"
set C2R_15_X32_ST="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRunStore"
set C2R_16_X64_ST="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"
set C2R_16_X32_ST="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore"

%printCmd% reg query %C2R_14_X64_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X64_ST% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X64_ST%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_14_X32_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X32_ST% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X32_ST%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X64_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X64_ST% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X64_ST%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X32_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X32_ST% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X32_ST%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X64_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X64_ST% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X64_ST%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X32_ST% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X32_ST% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X32_ST%
%PrintLog% echo ******************************************

set C2R_14_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\ClickToRun\Configuration"
set C2R_14_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\ClickToRun\Configuration"
set C2R_15_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
set C2R_15_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
set C2R_16_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
set C2R_16_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"

%printCmd% reg query %C2R_14_X64% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X64% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X64%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_14_X32% /v ProductReleaseIds
%PrintLog% reg query %C2R_14_X32% /v ProductReleaseIds && set c2r_14_Key=%C2R_14_X32%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X64% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X64% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X64%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_X32% /v ProductReleaseIds
%PrintLog% reg query %C2R_15_X32% /v ProductReleaseIds && set c2r_15_Key=%C2R_15_X32%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X64% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X64% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X64%
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_X32% /v ProductReleaseIds
%PrintLog% reg query %C2R_16_X32% /v ProductReleaseIds && set c2r_16_Key=%C2R_16_X32%
%PrintLog% echo ******************************************

set C2R_14_GUID_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\ClickToRun"
set C2R_14_GUID_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\ClickToRun"
set C2R_15_GUID_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
set C2R_15_GUID_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun"
set C2R_16_GUID_X64="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun"
set C2R_16_GUID_X32="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun"

%printCmd% reg query %C2R_14_GUID_X64% /v PackageGUID
%PrintLog% reg query %C2R_14_GUID_X64% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_14_GUID_X64% /v PackageGUID') do set guid_14=%%g)
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_14_GUID_X32% /v PackageGUID
%PrintLog% reg query %C2R_14_GUID_X32% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_14_GUID_X32% /v PackageGUID') do set guid_14=%%g)
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_GUID_X64% /v PackageGUID
%PrintLog% reg query %C2R_15_GUID_X64% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_15_GUID_X64% /v PackageGUID') do set guid_15=%%g)
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_15_GUID_X32% /v PackageGUID
%PrintLog% reg query %C2R_15_GUID_X32% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_15_GUID_X32% /v PackageGUID') do set guid_15=%%g)
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_GUID_X64% /v PackageGUID
%PrintLog% reg query %C2R_16_GUID_X64% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_16_GUID_X64% /v PackageGUID') do set guid_16=%%g)
%PrintLog% echo ******************************************

%printCmd% reg query %C2R_16_GUID_X32% /v PackageGUID
%PrintLog% reg query %C2R_16_GUID_X32% /v PackageGUID && (for /f "tokens=3" %%g in ('reg query %C2R_16_GUID_X32% /v PackageGUID') do set guid_16=%%g)
%PrintLog% echo ******************************************

set OSPP_HKLM=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set OSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set XSPP_HKLM_X32=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_HKLM_X64=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
call :CleanRegistryKeys

goto :eof

:convert

set tProduct=
set tYear=

set tProduct=%1
set tYear=%2

if /i '!tProduct!' EQU '' (
	echo product not exist .. wtf 
	goto :eof
)

if /i '!tYear!' EQU '' (
	echo product not exist .. wtf 
	goto :eof
)

set foundYear=
for %%g in (2010, 2013, 2016, 2019, 2021) do (if /i '!tYear!' EQU '%%g' (set foundYear=true))
if not defined foundYear (
	echo product not exist .. wtf
	goto :eof
)

%printCmd% dir /b licence\!tProduct!*!tYear!*.xrm-ms
(%PrintLog% dir /b licence\!tProduct!*!tYear!*.xrm-ms) || (
	%PrintLog% echo ******************************************
	echo product not exist .. wtf 
	goto :eof
)
%PrintLog% echo ******************************************

if not defined HideMode echo.
if not defined HideMode echo Convert exiting office to Office !tProduct! !tYear!
if not defined HideMode echo ................................................

set "ProductYear=!tYear!"
call :integrate !tProduct! !ProductYear!
for /f %%y in ('dir /b licence\!tProduct!*!tYear!*.xrm-ms') do call :inslic %%y !tProduct!

if defined Officec2r_v16 (
	call :UnInstallKeys 2016
	call :UnInstallKeys 2019
	call :UnInstallKeys 2021
) else (
	call :UnInstallKeys !ProductYear!
)

set LetWork2010=
if defined OfficeMsi_14			set LetWork2010=true
if '!ProductYear!' EQU '2010' 	set LetWork2010=true
		
if defined LetWork2010 (	
	if defined Office14_WOW (
		echo Import [Wow] Keys
		call :CmdWorker reg import Keys\!tProduct!\RegistrationWOW.reg
	) else (
		if /i '!xBit!' EQU '32' (
			echo Import [X32] Keys
			call :CmdWorker reg import Keys\!tProduct!\Registration32.reg
		)
		if /i '!xBit!' EQU '64' (
			echo Import [X32] Keys
			echo Import [X64] Keys
			call :CmdWorker reg import Keys\!tProduct!\Registration32.reg
			call :CmdWorker reg import Keys\!tProduct!\Registration64.reg
		)
	)
)

:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
>%tmpFile% call :volume_license_serial_list
for /f "tokens=1,2,3 delims=*" %%p in ('type %tmpFile%') do (
	if /i '%%q' EQU '!tYear!' (
		echo %%p | find /i "!tProduct!">nul && call :inpkey %%r
	)
)

:: convert other programs if have
set ProductsList=publisher,ProjectPro,ProjectStd,VisioStd,VisioPro,VisioPrem
for %%x in (!ProductsList!) do (

	set SelectedX=
	if defined OfficeMsi_14 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office14.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_15 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office15.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined OfficeMsi_16 (
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
		reg query "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"|>nul find /i "Office16.%%x" && (
			set "SelectedX=%%x"
		)
	)
	if defined Officec2r_v15 (
		echo !ProductReleaseIds! | >nul find /i "%%x" && (
			set "SelectedX=%%x"
		)
	)					
	if defined Officec2r_v16 (
		echo !ProductReleaseIds! | find /i "%%x">nul && (
			set "SelectedX=%%x"
			for %%m in (!ProductReleaseIds!) do (
				echo %%m | >nul find /i "%%x" && (
					set ProductYear=
					echo %%m | >nul find /i "2021" && set ProductYear=2021
					echo %%m | >nul find /i "2019" && set ProductYear=2019
					if not defined ProductYear 		  set ProductYear=2016
				)
			)
		)
	)
	if defined SelectedX (
		echo.
		echo Install Volume Licence serials for %%x !ProductYear!
		echo .................................................
		
		set LetWork2010=
		if defined OfficeMsi_14			set LetWork2010=true
		if '!ProductYear!' EQU '2010' 	set LetWork2010=true
		
		if defined LetWork2010 (
			if defined Office14_WOW (
				echo Import [Wow] Keys
				call :CmdWorker reg import Keys\%%x\RegistrationWOW.reg
			) else (
				if /i '!xBit!' EQU '32' (
					echo Import [X32] Keys
					call :CmdWorker reg import Keys\%%x\Registration32.reg
				)
				if /i '!xBit!' EQU '64' (
					echo Import [X32] Keys
					echo Import [X64] Keys
					call :CmdWorker reg import Keys\%%x\Registration32.reg
					call :CmdWorker reg import Keys\%%x\Registration64.reg
				)
			)
		)
		
		call :integrate %%x !ProductYear!
		for /f %%y in ('dir /b licence\%%x*!ProductYear!*.xrm-ms') do call :inslic  %%y %%x
		
		:: %%p ProjectPro %%q 2019 %%r B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
		>%tmpFile% call :volume_license_serial_list
		for /f "tokens=1,2,3 delims=*" %%p in ('type %tmpFile%') do (
			if /i '%%q' EQU '!ProductYear!' (
				echo %%p | find /i "%%x">nul && call :inpkey %%r
			)
		)
	)	
)

if defined xArgs (
	echo.
	call :Activate !ScriptMode! !ProductYear!
)
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

:volume_license_serial_list
echo Word*2010*HVHB3-C6FV7-KQX9W-YQG79-CRY7T
echo Word*2013*6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7
echo Word*2016*WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
echo word*2019*PBX3G-NWMT6-Q7XBW-PYJGG-WXD33
echo word*2021*TN8H9-M34D3-Y64V9-TR72V-X79KV
echo Excel*2010*H62QG-HXVKF-PP4HP-66KMR-CW9BM
echo Excel*2013*VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB
echo Excel*2016*9C2PK-NWTVB-JMPW8-BFT28-7FTBF
echo Excel*2019*TMJWT-YYNMB-3BKTF-644FC-RVXBD
echo Excel*2021*NWG3X-87C9K-TC7YY-BC2G7-G6RVC
echo PowerPoint*2010*RC8FX-88JRY-3PF7C-X8P67-P4VTT
echo PowerPoint*2013*4NT99-8RJFH-Q2VDH-KYG2C-4RD4F
echo PowerPoint*2016*J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
echo PowerPoint*2019*RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
echo PowerPoint*2021*TY7XF-NFRBR-KJ44C-G83KF-GX27K
echo ProPlus*2010*VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo ProPlus*2013*YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo ProPlus*2016*XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo ProPlus*2019*NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo ProPlus*2021*FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
echo VisioPro*2010*D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
echo VisioPro*2013*C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
echo VisioPro*2016*PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
echo VisioPro*2019*9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
echo VisioPro*2021*KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4
echo VisioStd*2010*767HD-QGMWX-8QTDB-9G3R2-KHFGJ
echo VisioStd*2013*J484Y-4NKBF-W2HMG-DBMJC-PGWR7
echo VisioStd*2016*7WHWN-4T7MP-G96JF-G33KR-W8GF4
echo VisioStd*2019*7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
echo VisioStd*2021*MJVNY-BYWPY-CWV6J-2RKRT-4M8QG
echo VisioPrem*2010*D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
echo Publisher*2010*BFK7F-9MYHM-V68C7-DRQ66-83YTP
echo Publisher*2013*PN2WF-29XG2-T9HJ7-JQPJR-FCXK4
echo Publisher*2016*F47MM-N3XJP-TQXJ9-BP99D-8K837
echo Publisher*2019*G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
echo Publisher*2021*2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ
echo ProjectPro*2010*YGX6F-PGV49-PGW3J-9BTGG-VHKC6
echo ProjectPro*2013*FN8TT-7WMH6-2D4X9-M337T-2342K
echo ProjectPro*2016*YG9NW-3K39V-2T3HJ-93F3Q-G83KT
echo ProjectPro*2019*B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
echo ProjectPro*2021*FTNWT-C6WBT-8HMGF-K9PRX-QV9H8
echo ProjectStd*2010*4HP3K-88W3F-W2K3D-6677X-F9PGB
echo ProjectStd*2013*6NTH3-CW976-3G3Y2-JK3TX-8QHTT
echo ProjectStd*2016*GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
echo ProjectStd*2019*C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
echo ProjectStd*2021*J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T
echo SkypeforBusiness*2016*869NQ-FJ69K-466HW-QYCP2-DDBV6
echo SkypeforBusiness*2019*NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
echo SkypeforBusiness*2021*HWCXN-K3WBT-WJBKY-R8BD9-XK29P
echo Standard*2010*V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo Standard*2013*KBKQT-2NMXY-JJWGP-M62JB-92CD4
echo Standard*2016*JNRGM-WHDWX-FJJG3-K47QV-DRTFM
echo Standard*2019*6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK
echo Standard*2021*KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3
echo mondo*2013*42QTK-RN8M7-J3C4G-BBGYM-88CYV
echo mondo*2016*HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
echo outlook*2010*7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ
echo outlook*2013*QPN8Q-BJBTJ-334K3-93TGY-2PMBT
echo outlook*2016*R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
echo outlook*2019*7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
echo outlook*2021*C9FM6-3N72F-HFJXB-TM3V9-T86R9
echo access*2010*V7Y44-9T38C-R2VJK-666HK-T7DDX
echo access*2013*NG2JY-H4JBT-HQXYP-78QH9-4JM2D
echo access*2016*GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
echo access*2019*9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT
echo access*2021*WM8YG-YNGDD-4JHDC-PG3F4-FC4T4
echo SmallBusBasics*2010*D6QFG-VBYP2-XQHM7-J97RH-VVRCK
echo OneNote*2016*DR92N-9HTF2-97XKM-XW2WJ-XW3J6
goto :eof

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

:: VBS Helpers
:: VBS Helpers
:: VBS Helpers

:Query

set pStart=%time: =0%

Rem Main Code Here
>nul 2>&1 del /q "%temp%\result"
if /i '%3' EQU '' (
	>"%temp%\result" cscript "tools\KmsHelper.vbs" "/QUERY_BASIC" %1 %2
) else (
	>"%temp%\result" cscript "tools\KmsHelper.vbs" "/QUERY_ADVENCED" %1 %2 %3
)

Rem Main Code Here

set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
 if /i '%3' EQU '' (
	%printCmd_v2% cscript "tools\KmsHelper.vbs" "/QUERY_BASIC" %1 %2
) else (
	%printCmd_v2% cscript "tools\KmsHelper.vbs" "/QUERY_ADVENCED" %1 %2 %3
)
%PrintLog% type "%temp%\result"
%PrintLog% echo ******************************************

goto :eof

:ACTIVATE_VBX

set pStart=%time: =0%

Rem Main Code Here
set "lastErr="
set "activationCMD=2>nul cscript "Tools\KmsHelper.vbs" "/ACTIVATE" "%1" "%2""
for /f "tokens=1,2 delims=: " %%x in ('"!activationCMD!"') do set "lastErr=%%y"
Rem Main Code Here

set pEnd=%time: =0%
call :CalcDiff %pStart% %pEnd%

%PrintLog% echo.
%PrintLog% echo Start :: !pStart! --- End :: !pEnd! --- Total !_tdiff!
%printCmd_v2% cscript "Tools\KmsHelper.vbs" "/ACTIVATE" "%1" "%2"
%PrintLog% echo Last Error !lastErr!
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
	copy /y "%~dp0\Tools\bin\%xOS%.dll" %_Hook% >nul 2>&1
	goto :skipsym
)
mklink %_Hook% "%~dp0\Tools\bin\%xOS%.dll" >nul 2>&1
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