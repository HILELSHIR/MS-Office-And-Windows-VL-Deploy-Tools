@cls
@echo off
>nul chcp 437

mode 110, 30
set LocalKms=true
set RemoveTeams=true
set cnfg="%temp%\tmp.xml"
set stpp="Tools\setup.exe"
set user_agent=Microsoft-Delivery-Optimization/10.0
set officecdn=officecdn.microsoft.com.edgesuite.net/pr
SETLOCAL EnableDelayedExpansion
title Local Installation Tool

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
)

set XSPP_HKLM_X32=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_HKLM_X64=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
set XSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform

:: basic validtion
call :cleanVariable

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

if defined RemoveTeams (
	1>nul 2>&1 REG ADD HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\common\officeupdate /v preventteamsinstall /t REG_DWORD /d 1 /f /reg:32
) else (
	1>nul 2>&1 REG ADD HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\common\officeupdate /v preventteamsinstall /t REG_DWORD /d 0 /f /reg:32
)

:: fix for running from mounted iso
set "fCD=%cd%"
if /i '%fCD:~-1%' EQU '\' (
	set "fCD=%fCD:~0,-1%")

:: Find Version Function ~ Start
echo.
echo Verify Version
if not exist "office\data" (
	echo.
	echo ~~ Folder Error ~~
	goto :End
)
set "version="
(1>nul 2>&1 dir /ad /b office\data\15.*) && for /f %%g in ('dir /ad /b office\data\15.*') do set "version=%%g"
(1>nul 2>&1 dir /ad /b office\data\16.*) && for /f %%g in ('dir /ad /b office\data\16.*') do set "version=%%g"
if not defined version (
	echo.
	echo Could not found office version
	goto :end
)
:: Find Version Function ~ End

rem St[1] default pharams
rem St[1] default language
set "userSelected="
call :ZeroValues 32 2021
call :LangPharser null
rem St[2] pharams - from system
rem St[2] language - from system
call :CheckSystemBit
call :Get-WinUserLanguageList_Warper
call :CheckSystemLanguage
call :UpdateSystemSettings
rem St[3] Get Info From Directory
call :GetFolderInfo
call :UpdateFolderInfo
rem St[4] language - user prompt
set "xVal="&set "xVal=%*"
if defined xVal call :InputPharser %xVal%
rem St[5] read from config file
call :UpdateBuildInfo
rem done checking
call :UpdatePID

:: get build number
set "BuildNumber="
set "Windows_7_Or_Earlier="
call :query "select * from Win32_OperatingSystem" "buildnumber" 
for /f "tokens=* skip=3" %%g in ('type "%temp%\result"') do set "BuildNumber=%%g"
set "BuildNumber=!BuildNumber: =!"

if /i !buildnumber! LEQ 7601 set "Windows_7_Or_Earlier=true"
if defined Windows_7_Or_Earlier (
	if /i '!ProductYear!' NEQ '2016' (
		if /i '!ProductYear!' NEQ '2013' (
			echo.
			echo Can't install office 2019 - 2021 on windows 7
			goto :end
		)
	)
	
	if /i !version! GTR 16.0.12527.22021 (
		echo.
		echo Can't install office 2016 V!version! On Windows 7
		goto :end
	)
)

if !buildnumber! LSS 2600 (
	echo.
	echo ERROR ### System not supported
	echo.
	pause
	exit /b
)

:: sBit is system Type
:: xBit is desire Type
if /i !xBit! GTR !sBit! (
	echo.
	echo you try installing Office x64 on x32 System
	goto :end
)

if not defined userSelected (
	if not defined lYear (
		echo:
		echo Select Product Year
		echo:
		echo :: 1 - 2013
		echo :: 2 - 2016
		echo :: 3 - 2019
		echo :: 4 - 2021
		echo :: E - Exit
		echo:
		CHOICE /C 1234E /M "Select build Year To Continue, Press E to Exit ::"
		if !errorlevel! EQU 1 call :updateYear 2013
		if !errorlevel! EQU 2 call :updateYear 2016
		if !errorlevel! EQU 3 call :updateYear 2019
		if !errorlevel! EQU 4 call :updateYear 2021
		if !errorlevel! EQU 5 (pause & exit /b)
		call :UpdatePID
	)
)

:: office install
if not defined setupMode (
	echo:
	echo :: Select Build options
	echo:
	echo :: 1 - Lite
	echo :: 2 - Full
	echo :: 3 - VIP
	echo :: 4 - 365
	echo :: 5 - Custom
	echo :: 6 - ALL
	echo :: E - Exit
	echo:
	CHOICE /C 123456E /M "Select build type To Continue, Press E to Exit ::"
	
	if !errorlevel! EQU 1 set /a "setupMode=1"
	if !errorlevel! EQU 2 (
		echo.
		set /a "setupMode=2"
		set "selected_Products="
		
		set "xzzyz5="
		set /p xzzyz5=Install Word ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,word"
			) else (
				set "selected_Products=word"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Excel ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,excel"
			) else (
				set "selected_Products=excel"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install PowerPoint ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,powerpoint"
			) else (
				set "selected_Products=powerpoint"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Access ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,access"
			) else (
				set "selected_Products=access"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Outlook ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,outlook"
			) else (
				set "selected_Products=outlook"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Publisher ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,publisher"
			) else (
				set "selected_Products=publisher"
			)
		)
		
		if /i '!ProductYear!' NEQ '2013' (
			set "xzzyz5="
			set /p xzzyz5=Install Skype For Business ^? [Enter to Install, Any key to Exclude] ^>
			if defined xzzyz5 (
				if defined selected_Products (
					set "selected_Products=!selected_Products!,skypeforbusiness"
				) else (
					set "selected_Products=skypeforbusiness"
				)
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install OneNote ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,onenote"
			) else (
				set "selected_Products=onenote"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Lync ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,lync"
			) else (
				set "selected_Products=lync"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Teams ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,teams"
			) else (
				set "selected_Products=teams"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Groove ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,groove"
			) else (
				set "selected_Products=groove"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install OneDrive ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,onedrive"
			) else (
				set "selected_Products=onedrive"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Bing ^? [Enter to Install, Any key to Exclude] ^>
		if defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,bing"
			) else (
				set "selected_Products=bing"
			)
		)
	)
	if !errorlevel! EQU 3 set /a "setupMode=3"
	if !errorlevel! EQU 4 (
		set /a "setupMode=4"
		call :updateYear 2016
		call :UpdatePID
	)
	if !errorlevel! EQU 5 (
		echo.
		set /a "setupMode=5"
		set "selected_Products="
		
		set "xzzyz5="
		set /p xzzyz5=Install Word ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,Word"
			) else (
				set "selected_Products=Word"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Excel ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,Excel"
			) else (
				set "selected_Products=Excel"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install PowerPoint ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,PowerPoint"
			) else (
				set "selected_Products=PowerPoint"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Access ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,Access"
			) else (
				set "selected_Products=Access"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Outlook ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,Outlook"
			) else (
				set "selected_Products=Outlook"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Publisher ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,Publisher"
			) else (
				set "selected_Products=Publisher"
			)
		)
		
		if /i '!ProductYear!' NEQ '2013' (
			set "xzzyz5="
			set /p xzzyz5=Install Skype For Business ^? [Enter to Install, Any key to Exclude] ^>
			if not defined xzzyz5 (
				if defined selected_Products (
					set "selected_Products=!selected_Products!,SkypeForBusiness"
				) else (
					set "selected_Products=SkypeForBusiness"
				)
			)
		)
		
		if /i '!ProductYear!' NEQ '2013' (
			if /i '!ProductYear!' NEQ '2019' (
				set "xzzyz5="
				set /p xzzyz5=Install OneNote ^? [Enter to Install, Any key to Exclude] ^>
				if not defined xzzyz5 (
					if defined selected_Products (
						set "selected_Products=!selected_Products!,OneNote"
					) else (
						set "selected_Products=OneNote"
					)
				)
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Visio Pro ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,VisioPro"
			) else (
				set "selected_Products=VisioPro"
			)
		)
		
		set "xzzyz5="
		set /p xzzyz5=Install Project Pro ^? [Enter to Install, Any key to Exclude] ^>
		if not defined xzzyz5 (
			if defined selected_Products (
				set "selected_Products=!selected_Products!,ProjectPro"
			) else (
				set "selected_Products=ProjectPro"
			)
		)
		
		if not defined selected_Products (
			echo.
			echo ERROR ### No Products Selected
			echo.
			timeout 5
			exit /b
		)
	)
	if !errorlevel! EQU 6 set /a "setupMode=6"
	if !errorlevel! EQU 7 (pause & exit /b)
)

if defined setupMode (
	if '!setupMode!' EQU '5' (
		if not defined selected_Products (
			echo.
			echo ERROR ### No Products Selected
			echo.
			timeout 5
			exit /b
		)
	)
)

:: office install
echo.
if defined setupmode (

	echo !setupMode!|>nul find /i "1" && echo Install Word,Excel,Powerpoint !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
	
	echo !setupMode!|>nul find /i "2" && (
		if /i '!ProductYear!' EQU '2013' (
			echo Install Office professional plus !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
		) else if /i '!ProductYear!' EQU '2016' (
			echo Install Office professional plus !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
		) else (
			echo Install Office professional plus !ProductYear! Perpetual VL %langCd:~0,2% x!xBit! v!version!
		)
	)
	
	echo !setupMode!|>nul find /i "3" && echo Install ProjectPro,visioPro !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
	echo !setupMode!|>nul find /i "4" && echo Install Microsoft 365 Business %langCd:~0,2% x!xBit! v!version!
	echo !setupMode!|>nul find /i "5" && echo Install !selected_Products! !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
	echo !setupMode!|>nul find /i "6" && echo Install ProPlus,ProjectPro,visioPro !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
)

if not defined setupmode (
	if /i '!ProductYear!' EQU '2013' (
		echo Install Office professional plus !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
	) else if /i '!ProductYear!' EQU '2016' (
		echo Install Office professional plus !ProductYear! VL %langCd:~0,2% x!xBit! v!version!
	) else (
		echo Install Office professional plus !ProductYear! Perpetual VL %langCd:~0,2% x!xBit! v!version!
	)
)

:: check for Old Office Licence
if /i "!version:~0,2!" EQU "16" (
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)
if /i "!version:~0,2!" EQU "15" (
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRunStore"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
)
if defined key (
	for /f "tokens=2 delims=_" %%g in ('reg query !key! /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds=%%g
	set ProductReleaseIds="!ProductReleaseIds:~6!"
)

:: odt ~ Nethood
REM call :usingODT

:: OfficeClickToRun ~ Nethood
call :usingPRSys

:: check for New Office Licence
if not defined key (
	if /i "!version:~0,2!" EQU "16" (
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\propertyBag"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\ClickToRunStore"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	)
	if /i "!version:~0,2!" EQU "15" (
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRunStore"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"
		1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" /v ProductReleaseIds && set key="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
	)
	
	if not defined key (
		
		rem Fail on :: Key Never Created
		rem Fail on :: Key Never Created
		rem Fail on :: Key Never Created
	
		echo.
		echo Installation Failed
		goto :end
	)
)

if defined ProductReleaseIds (
	for /f "tokens=2 delims=_" %%g in ('reg query !key! /v ProductReleaseIds ^| find /i "ProductReleaseIds"') do set ProductReleaseIds_=%%g
	set ProductReleaseIds_="!ProductReleaseIds_:~6!"
	if /i !ProductReleaseIds! EQU !ProductReleaseIds_! (
		
		if not defined OkStatus (
			rem FAIL on :: Key exist, but didnt change
			rem FAIL on :: Key exist, but didnt change
			rem FAIL on :: Key exist, but didnt change
		
			echo.
			echo Installation Failed
			goto :end
		)
	)
	
	rem Continue if Key was Changes during installation
	rem Continue if Key was Changes during installation
	rem Continue if Key was Changes during installation
)

rem Continue if New key Created
rem Continue if New key Created
rem Continue if New key Created

::licence Install
call :LicenceInstall !sBit!
if defined LocalKms 		call :activate_Local
if not defined LocalKms 	call :activate_Online

:end

rem Windows_Addict Not genuine fix
1>nul 2>nul reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
1>nul 2>nul reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"
1>nul 2>nul reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "%IP_ADDRESS%"

echo.
echo please wait 5 seconds..
SETLOCAL DisableDelayedExpansion
timeout /t 6 /NOBREAK
goto :eof

:cleanVariable
set vars=xVal, xBit, sBit, sLng, lLngID, lBit, lLng, lYear, equal, multi, stream, langId, langCd, version, cscript, platform, file_list, iso_file, ProductId, SysLanCD, SysLanIdHex, ProductYear, SupportedBit, VersionReleaseApi, SupportedbuildYear, key, ProductReleaseIds, ProductReleaseIds_
for %%v in (%vars%) do (set %%v=)
goto :eof

:stpDownload
set stpp="%temp%\setup.exe"
1>nul 2>&1 tools\wget --quiet --tries=3 --retry-connrefused --output-document=%stpp% "https://officecdn.microsoft.com/pr/wsus/setup.exe"
goto :eof

:UpdateBuildInfo
set "SupportedBit=32, 64, Multi"
set "SupportedbuildYear=2013, 2016, 2019, 2021"
if not exist BuildInfo.ini goto :eof
for /f "tokens=1,* delims==" %%g in ('type BuildInfo.ini') do call :BuildInfoPharser %%g %%h
goto :eof

:BuildInfoPharser
if /i '%1' EQU 'xBit' (
	for %%g in (%SupportedBit%) do (
		if /i '%%g' EQU '%2' call :updateBit %%g
	)
)
if /i '%1' EQU 'ProductYear' ( 
	for %%g in (%SupportedbuildYear%) do (
		if /i '%%g' EQU '%2' (
			set "userSelected=***"
			call :updateYear %%g
		)
	)
)
if /i '%1' EQU 'Version' (
	set "version=%2"
	call :UpdatePID
)
if /i '%1' EQU 'LanguageName' (
	call :BuildInfoPharserLanguageExtender %*
)
if /i '%1' EQU 'LanguageCode' (
	call :BuildInfoPharserLanguageExtenderX %*
)
if /i '%1' EQU 'Mode' (
	for %%$ in (1,2,3,4,5,6) do (
		if /i '%%$' EQU '%2' (
			set /a "setupMode=%2"
		)
	)
	
	if defined setupMode (
		if !setupMode! EQU 4 (
			call :updateYear 2016
			call :UpdatePID
		)
	)
)
if /i '%1' EQU 'IncludeProducts' (
	set "selected_Products="
	set "productList=%*"
	set "productList=!productList:~16!"
	
	for %%$ in (word,excel,powerpoint,access,outlook,publisher,skypeforbusiness,visiopro,projectpro) do (
		echo '!productList!'|>nul find /i "%%$" && (
			if defined selected_Products (
				echo "!selected_Products!" |>nul find /i "%%$" || (
					set "selected_Products=!selected_Products!,%%$")
			) else (
				set "selected_Products=%%$"
			)
		)
	)
)

if /i '%1' EQU 'ExcludeProducts' (
	set "selected_Products="
	set "productList=%*"
	set "productList=!productList:~16!"
	
	rem Must be LowerCase LOL
	for %%$ in (word,excel,powerpoint,access,outlook,publisher,skypeforbusiness,lync,teams,groove,onedrive,bing,onenote) do (
		echo '!productList!'|>nul find /i "%%$" && (
			if defined selected_Products (
				echo "!selected_Products!" |>nul find /i "%%$" || (
					set "selected_Products=!selected_Products!,%%$")
			) else (
				set "selected_Products=%%$"
			)
		)
	)
)
goto :eof

:BuildInfoPharserLanguageExtender
set "value=%*"
set "value=%value:~13%"
call :LangPharser !value!
goto :eof

:BuildInfoPharserLanguageExtenderX
set "value=%*"
set "value=%value:~13%"
call :LangPharserX !value!
goto :eof

:updateYear
if defined Ignore (
	set "Ignore="
	goto :eof
)
set ProductYear=%*
goto :eof

:updateBit
set xBit=%*
if /i '%*' EQU 'Multi' (
	set multi=true
	set xBit=!sBit!
)
goto :eof

:ZeroValues
call :updateBit %1
call :updateYear %2
goto :eof

:LicenceInstall
if /i '%1' NEQ '' (
	
	:: second loop
	if not defined cscript (
		if /i '%1' EQU '64' set cscript="%windir%\SysWOW64\cscript.exe"
		if /i '%1' EQU '32' set cscript="%windir%\system32\cscript.exe"
		echo Install Office Volume Licence
		call :LicenceInstall %1 FuckNoMore
		goto :eof
	)
	
	:: third loop ~ END
	if /i '%2' NEQ '' (
	
		if defined setupmode (
			echo !setupMode!|>nul find /i "1" && (
				for %%g in (word, excel, powerpoint) do (
			
					for /f %%i in ('dir /b licence\%%g!ProductYear!*.xrm-ms') do (
						1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
					)
					
					:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
						if /i '%%x' EQU '%%g' (
							if /i '%%y' EQU '!ProductYear!' (
								1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
							)
						)
					)
				)
			)
			
			echo !setupMode!|>nul find /i "2" && (
				for /f %%i in ('dir /b licence\ProPlus!ProductYear!*.xrm-ms') do 1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
				
				:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
					if /i '%%x' EQU 'ProPlus' (
						if /i '%%y' EQU '!ProductYear!' (
							1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
						)
					)
				)
			)
			
			echo !setupMode!|>nul find /i "3" && (
				for %%g in (ProjectPro,visioPro) do (
			
					for /f %%i in ('dir /b licence\%%g!ProductYear!*.xrm-ms') do (
						1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
					)
					
					:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
						if /i '%%x' EQU '%%g' (
							if /i '%%y' EQU '!ProductYear!' (
								1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
							)
						)
					)
				)
			)
			
			echo !setupMode!|>nul find /i "4" && (
				for /f %%i in ('dir /b licence\Mondo!ProductYear!*.xrm-ms') do 1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
				
				:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
					if /i '%%x' EQU 'Mondo' (
						if /i '%%y' EQU '!ProductYear!' (
							1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
						)
					)
				)
			)
			
			echo !setupMode!|>nul find /i "5" && (
				for %%g in (!selected_Products!) do (
					(echo "%%g" |>nul find /i "OneNote") && (
						for /f %%i in ('dir /b licence\OneNote2016*.xrm-ms') do (
							1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
						)
					) || (
						for /f %%i in ('dir /b licence\%%g!ProductYear!*.xrm-ms') do (
							1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
						)
					)
					
					:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
						if /i '%%x' EQU '%%g' (
							if /i '%%y' EQU '!ProductYear!' (
								1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
							)
						)
					)
				)
			)
			
			echo !setupMode!|>nul find /i "6" && (
				for %%g in (ProPlus,ProjectPro,visioPro) do (
			
					for /f %%i in ('dir /b licence\%%g!ProductYear!*.xrm-ms') do (
						1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
					)
					
					:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
					>"%temp%\tmp" call :volume_license_serial_list
					for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
						if /i '%%x' EQU '%%g' (
							if /i '%%y' EQU '!ProductYear!' (
								1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
							)
						)
					)
				)
			)
		
			goto :eof
		)
		
		for /f %%i in ('dir /b licence\ProPlus!ProductYear!*.xrm-ms') do 1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inslic:"Licence\%%i"
		
		:: %%x ProjectPro %%y 2019 %%z B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
			>"%temp%\tmp" call :volume_license_serial_list
			for /f "tokens=1,2,3 delims=*" %%x in ('type "%temp%\tmp"') do (
			if /i '%%x' EQU 'ProPlus' (
				if /i '%%y' EQU '!ProductYear!' (
					1>nul 2>&1 %cscript% Tools\x%1\ospp.vbs /inpkey:%%z
				)
			)
		)
		
		goto :eof
	)
)

:: first loop
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 call :LicenceInstall 32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 call :LicenceInstall 64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	call :LicenceInstall 64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	call :LicenceInstall 64
goto :eof

:UpdateSystemSettings
call :ZeroValues %sBit% 2021
call :LangPharser %sLng%
goto :eof

:CheckSystemBit
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set sBit=32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set sBit=64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set sBit=64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set sBit=64
goto :eof

:UpdateLanguage
set langId=%1
set langCd=%2
goto :eof

:LangPharser
:: %%g=English %%h=1033 %%i=en-us %%j:0409
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
	call :compareXY "%*" "%%g"
	if defined equal (
		call :UpdateLanguage %%h %%i
		goto :LangPharser_
	)
)
:LangPharser_
goto :eof

:LangPharserX
:: %%g=English %%h=1033 %%i=en-us %%j:0409
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
	call :compareXY "%*" "%%i"
	if defined equal (
		call :UpdateLanguage %%h %%i
		goto :LangPharserX_
	)
)
:LangPharserX_
goto :eof

:InputPharser
set pharse=%*

:: year is '%pharse:~0,4%'
:: bit is '%pharse:~5,2%'
:: lang is '%pharse:~8%'

for %%g in (1,2,3,4,6) do (
	if /i "%pharse:~0,1%" EQU "%%g" (
		if /i "%pharse:~0,2%" NEQ "20" (
			goto :InputPharser_S
			exit /b
		)
	)
)

call :compareXY "%pharse:~0,4%" ""
if not defined equal (
	for %%g in (2013, 2016, 2019, 2021) do (
		echo "%pharse:~0,4%"|>nul find /i "%%g" && (
			set "userSelected=***"
			call :updateYear %%g
		)
	)
)

call :compareXY "%pharse:~5,5%" "multi"
if defined equal (
	call :updateBit multi
	goto :InputPharser_
)

call :compareXY "%pharse:~5,2%" ""
if not defined equal (
	for %%g in (32, 64) do (
		echo "%pharse:~5,2%"|>nul find /i "%%g" && call :updateBit %%g
	)
)

:InputPharser_
if defined multi (
	call :compareXY "%pharse:~11%" ""
	if not defined equal (
		call :LangPharser %pharse:~11%
	)
)

if not defined multi (
	call :compareXY "%pharse:~8%" ""
	if not defined equal (
		call :LangPharser %pharse:~8%
	)
)

goto :eof

:InputPharser_S
set pharse=%*

:: type is '%pharse:~0,1%'
:: year is '%pharse:~2,4%'
:: bit is '%pharse:~8,2%'
:: lang is '%pharse:~8%'

for %%g in (1,2,3,4,6) do (
	if /i "%pharse:~0,1%" EQU "%%g" (
		set /a "setupMode=%%g"
	)
)

echo !setupMode!|>nul find /i "4" && (
	set "userSelected=***"
	call :updateYear 2016
	call :UpdatePID
	call :compareXY "%pharse:~2%" ""
	if not defined equal (
		call :LangPharser %pharse:~2%
	)
	goto :eof
)

call :compareXY "%pharse:~2,4%" ""
if not defined equal (
	for %%g in (2013, 2016, 2019, 2021) do (
		echo "%pharse:~2,4%"|>nul find /i "%%g" && (
			set "userSelected=***"
			call :updateYear %%g
		)
	)
)

call :compareXY "%pharse:~7,5%" "multi"
if defined equal (
	call :updateBit multi
	goto :InputPharser__S
)

call :compareXY "%pharse:~7,2%" ""
if not defined equal (
	for %%g in (32, 64) do (
		echo "%pharse:~7,2%"|>nul find /i "%%g" && call :updateBit %%g
	)
)

:InputPharser__S
if defined multi (
	call :compareXY "%pharse:~13%" ""
	if not defined equal (
		call :LangPharser %pharse:~13%
	)
)

if not defined multi (
	call :compareXY "%pharse:~10%" ""
	if not defined equal (
		call :LangPharser %pharse:~10%
	)
)

goto :eof

:compareXY
set equal=
if /i %1 EQU %2 set equal=true
goto :eof

:CheckSystemLanguage
set var=&set var=%*
if not defined var (

	:: Using HKCR :: PreferredUILanguages Value
	1>nul 2>&1 reg query "HKEY_CURRENT_USER\Control Panel\Desktop" /v PreferredUILanguages && (
		REM echo Using HKCR :: PreferredUILanguages Value
		for /f "tokens=1,3" %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\Desktop" /v PreferredUILanguages') do (
			if /i '%%g' EQU 'PreferredUILanguages' call :CheckSystemLanguage %%h
		)
		goto :eof
	)
	
	:: Using HKLM:: PreferredUILanguages Value
	1>nul 2>&1 reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\MUI\Settings" /v PreferredUILanguages && (
		REM echo Using HKLM:: PreferredUILanguages Value
		for /f "tokens=1,3" %%g in ('reg query "HKEY_LOCAL_MACHINE\SYSTEM\SYSTEM\CurrentControlSet\Control\MUI\Settings" /v PreferredUILanguages') do (
			if /i '%%g' EQU 'PreferredUILanguages' call :CheckSystemLanguage %%h
		)
		goto :eof
	)
	
	:: using Get-WinUserLanguageList Function Cmd Warper
	if defined SysLanCD (
		call :CheckSystemLanguage %SysLanCD%
		goto :eof
	)
	
	:: using dism :: get-intl
	REM echo using dism :: get-intl
	for /f "tokens=4,6" %%g  in ('dism /online /get-intl') do (
		if /i '%%g' EQU 'LanguageName' call :CheckSystemLanguage %%h
	)
	goto :eof
)
if defined var (
		
	:: %%g=English %%h=1033 %%i=en-us %%j:0409
	>"%temp%\tmp" call :Language_List
	for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
		if /i '%var%' EQU '%%i' (
			set sLng=%%g
			goto :CheckSystemLanguage_
		)
	)
)
:CheckSystemLanguage_
goto :eof

:Get-WinUserLanguageList_Warper
call :Get-WinUserLanguageList
if defined SysLanIdHex call :convertLanHexToDec
goto :eof
:Get-WinUserLanguageList
set xVal=
set SysLanCD=
set SysLanIdHex=
1>nul 2>&1 reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile" /v Languages || goto :eof
for /f "tokens=3 delims= " %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile" /v Languages ^| find /i "REG_MULTI_SZ"') do set xVal=%%g
if defined xVal 		(for /f "tokens=1 delims=\0" %%g in ('echo !xVal!') do set SysLanIdHex=%%g)
if defined SysLanIdHex 	(for /f "tokens=1 delims= " %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile\!SysLanIdHex!" ^| find /i "000"') do set SysLanIdHex=%%g)
if defined SysLanIdHex 	(for /f "tokens=1 delims=:" %%g in ('echo !SysLanIdHex!') do set SysLanIdHex=%%g)
goto :eof
:convertLanHexToDec
:: %%g=English %%h=1033 %%i=en-us %%j:0409
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3,4 delims=*" %%g in ('type "%temp%\tmp"') do (
	if 	/i '%%j' EQU '!SysLanIdHex!' (
		set SysLanCD=%%i
		goto :convertLanHexToDec_
	)	
)
:convertLanHexToDec_
goto :eof

rem Overview of deploying languages for Microsoft 365 Apps
rem https://docs.microsoft.com/en-us/deployoffice/overview-deploying-languages-microsoft-365-apps\

rem 2.1.1906 Part 4 Section 7.6.2.39, LCID (Locale ID)
rem https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a

rem Language identifiers and OptionState Id values in Office 2016 - Deploy Office | Microsoft Docs
rem https://docs.microsoft.com/en-us/deployoffice/office2016/language-identifiers-and-optionstate-id-values-in-office-2016

rem Decimal to Hexadecimal converter -- Hex signed 2's complement
rem https://www.rapidtables.com/convert/number/decimal-to-hex.html

rem 2.2 LCID Structure
rem https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/63d3d639-7fd2-4afb-abbe-0d5b5551eef8

rem Hexadecimal to Decimal converter
rem https://www.rapidtables.com/convert/number/hex-to-decimal.html

rem Windows Locale Codes - Sortable list
rem https://www.science.co.il/language/Locale-codes.php

:Language_List
echo Afrikaans*1078*af-za*0436
echo Albanian*1052*sq-al*041c
echo Amharic*1118*am-et*045e
echo Arabic*1025*ar-sa*0401
echo Armenian*1067*hy-am*042b
echo Assamese*1101*as-in*044d
echo Azerbaijani Latin*1068*az-latn-az*042c
echo Bangla Bangladesh*2117*bn-bd*0845
echo Bangla Bengali India*1093*bn-in*0445
echo Basque Basque*1069*eu-es*042d
echo Belarusian*1059*be-by*0423
echo Bosnian*5146*bs-latn-ba*0141a
echo Bulgarian*1026*bg-bg*0402
echo Catalan Valencia*2051*ca-es-valencia*0803
echo Catalan*1027*ca-es*0403
echo Chinese Simplified*2052*zh-cn*0804
echo Chinese Traditional*1028*zh-tw*0404
echo Croatian*1050*hr-hr*041a
echo Czech*1029*cs-cz*0405
echo Danish*1030*da-dk*0406
echo Dari*1164*prs-af*048c
echo Dutch*1043*nl-nl*0413
echo English UK*2057*en-GB*0809
echo English*1033*en-us*0409
echo Estonian*1061*et-ee*0425
echo Filipino*1124*fil-ph*0464
echo Finnish*1035*fi-fi*040b
echo French Canada*3084*fr-CA*0C0C
echo French*1036*fr-fr*040c
echo Galician*1110*gl-es*0456
echo Georgian*1079*ka-ge*0437
echo German*1031*de-de*0407
echo Greek*1032*el-gr*0408
echo Gujarati*1095*gu-in*0447
echo Hausa Nigeria*1128*ha-Latn-NG*0468
echo Hebrew*1037*he-il*040d
echo Hindi*1081*hi-in*0439
echo Hungarian*1038*hu-hu*040e
echo Icelandic*1039*is-is*040f
echo Igbo*1136*ig-NG*0470
echo Indonesian*1057*id-id*0421
echo Irish*2108*ga-ie*083c
echo Italian*1040*it-it*0410
echo Japanese*1041*ja-jp*0411
echo Kannada*1099*kn-in*044b
echo Kazakh*1087*kk-kz*043f
echo Khmer*1107*km-kh*0453
echo KiSwahili*1089*sw-ke*0441
echo Konkani*1111*kok-in*0457
echo Korean*1042*ko-kr*0412
echo Kyrgyz*1088*ky-kg*0440
echo Latvian*1062*lv-lv*0426
echo Lithuanian*1063*lt-lt*0427
echo Luxembourgish*1134*lb-lu*046e
echo Macedonian*1071*mk-mk*042f
echo Malay Latin*1086*ms-my*043e
echo Malayalam*1100*ml-in*044c
echo Maltese*1082*mt-mt*043a
echo Maori*1153*mi-nz*0481
echo Marathi*1102*mr-in*044e
echo Mongolian*1104*mn-mn*0450
echo Nepali*1121*ne-np*0461
echo Norwedian Nynorsk*2068*nn-no*0814
echo Norwegian Bokmal*1044*nb-no*0414
echo Null*1033*en-us*0409
echo Odia*1096*or-in*0448
echo Pashto*1123*ps-AF*0463
echo Persian*1065*fa-ir*0429
echo Polish*1045*pl-pl*0415
echo Portuguese Brazilian*1046*pt-br*0416
echo Portuguese Portugal*2070*pt-pt*0816
echo Punjabi Gurmukhi*1094*pa-in*0446
echo Quechua*3179*quz-pe*0c6b
echo Romanian*1048*ro-ro*0418
echo Romansh*1047*rm-CH*0417
echo Russian*1049*ru-ru*0419
echo Setswana*1074*tn-ZA*0432
echo Scottish Gaelic*1169*gd-gb*0491
echo Serbian Bosnia*7194*sr-cyrl-ba*01c1a
echo Serbian Serbia*10266*sr-cyrl-rs*0281a
echo Serbian*9242*sr-latn-rs*0241a
echo Sesotho sa Leboa*1132*nso-ZA*046C
echo Sindhi Arabic*2137*sd-arab-pk*0859
echo Sinhala*1115*si-lk*045b
echo Slovak*1051*sk-sk*041b
echo Slovenian*1060*sl-si*0424
echo Spanish*3082*es-es*0c0a
echo Spanish Mexico*2058*es-MX*080A
echo Swedish*1053*sv-se*041d
echo Tamil*1097*ta-in*0449
echo Tatar Cyrillic*1092*tt-ru*0444
echo Telugu*1098*te-in*044a
echo Thai*1054*th-th*041e
echo Turkish*1055*tr-tr*041f
echo Turkmen*1090*tk-tm*0442
echo Ukrainian*1058*uk-ua*0422
echo Urdu*1056*ur-pk*0420
echo Uyghur*1152*ug-cn*0480
echo Uzbek*1091*uz-latn-uz*0443
echo Vietnamese*1066*vi-vn*042a
echo Welsh*1106*cy-gb*0452
echo Wolof*1160*wo-SN*0488
echo Yoruba*1130*yo-NG*046A
echo isiXhosa*1076*xh-ZA*0434
echo isiZulu*1077*zu-ZA*0435
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
echo mondo*2010*7TC2V-WXF6P-TD7RT-BQRXR-B8K32
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

:activate_Local
call :StartKMSActivation
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /sethst:!KMSHostIP!
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /setprt:!KMSPort!
(%cscript% Tools\x!xBit!\ospp.vbs /act | find /i "Product activation successful">nul) && (echo Product activation succeeded) || (echo Product activation failed)
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /remhst
call :StopKMSActivation
goto :eof

:activate_Online
echo looking for active servers
>"%temp%\tmp" call :Kms_Servers_List
for /f "tokens=*" %%g in ('type "%temp%\tmp"') do (
	1>nul 2>&1 tools\tcping -4 -n 1 -g 1 -w 0.5 -i 0.5 %%g 1688 && (
		set KmsServer=%%g
		goto :activate_
	)
)
echo.
echo didnt found any online kms server
goto :eof
:activate_
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /sethst:!KmsServer!
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /setprt:1688
echo Winner Winner Chicken dinner
(%cscript% Tools\x!xBit!\ospp.vbs /act | find /i "Product activation successful">nul) && (echo Product activation succeeded) || (echo Product activation failed)
1>nul 2>&1 %cscript% Tools\x!xBit!\ospp.vbs /remhst
goto :eof

:: Actual diffrents from scripts Moved here
:: its more elegent way 

:BuildConfigartionFile

if /i '!ProductYear!' EQU '2013' (
	1>nul 2>&1 del /q %cnfg%
	goto :eof
)

if /i '!ProductYear!' EQU '2016' (
	echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
	>%cnfg%  echo ^<Configuration^>
	>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
	>>%cnfg% echo     ^<Product ID^=^"ProPlusRetail^"^>
	>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
	>>%cnfg% echo     ^<^/Product^>
	>>%cnfg% echo ^<^/Add^>
	>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
	>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
	>>%cnfg% echo ^<^/Configuration^>
) else (
	echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
	>%cnfg%  echo ^<Configuration^>
	>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"PerpetualVL!ProductYear!^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
	>>%cnfg% echo     ^<Product ID^=^"ProPlus!ProductYear!Volume^"^>
	>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
	>>%cnfg% echo     ^<^/Product^>
	>>%cnfg% echo ^<^/Add^>
	>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
	>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL!ProductYear!^" ^/^>
	>>%cnfg% echo ^<^/Configuration^>
)

if defined setupmode (

	echo !setupMode!|>nul find /i "1" && (
		if /i '!ProductYear!' EQU '2016' (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"WordRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"ExcelRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"PowerPointRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		) else (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"PerpetualVL!ProductYear!^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"Word!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"Excel!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"PowerPoint!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"OneDrive^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL!ProductYear!^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		)
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "2" && (
		if /i '!ProductYear!' EQU '2016' (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProPlusRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		) else (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"PerpetualVL!ProductYear!^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProPlus!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL!ProductYear!^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		)
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "3" && (
		if /i '!ProductYear!' EQU '2016' (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProjectProRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"VisioProRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		) else (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"PerpetualVL!ProductYear!^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProjectPro!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"VisioPro!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL!ProductYear!^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		)
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "4" && (
		echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
		>%cnfg%  echo ^<Configuration^>
		>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
		>>%cnfg% echo     ^<Product ID^=^"O365BusinessRetail^"^>
		>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
		>>%cnfg% echo     ^<^/Product^>
		>>%cnfg% echo ^<^/Add^>
		>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
		>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
		>>%cnfg% echo ^<^/Configuration^>
	)
	
	echo !setupMode!|>nul find /i "6" && (
		if /i '!ProductYear!' EQU '2016' (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"Current^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProPlusRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"ProjectProRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"VisioProRetail^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"Current^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		) else (
			echo Build Configuration File for Office !ProductYear! x!xBit! !langCd!
			>%cnfg%  echo ^<Configuration^>
			>>%cnfg% echo ^<Add OfficeClientEdition^=^"!xBit!^" Channel^=^"PerpetualVL!ProductYear!^" Version^=^"!version!^" SourcePath^=^"%fCD%^"^>
			>>%cnfg% echo     ^<Product ID^=^"ProPlus!ProductYear!Volume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"ProjectProVolume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo     ^<Product ID^=^"VisioProVolume^"^>
			>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
			>>%cnfg% echo       ^<ExcludeApp ID^=^"Groove^" ^/^>
			>>%cnfg% echo     ^<^/Product^>
			>>%cnfg% echo ^<^/Add^>
			>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
			>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL!ProductYear!^" ^/^>
			>>%cnfg% echo ^<^/Configuration^>
		)
		goto :eof
	)
)
goto :eof

:usingODT
if not exist %stpp% (
	echo.
	echo Downloading Setup file
	call :stpDownload
)
call :BuildConfigartionFile
if not exist %cnfg% goto :eof
%stpp% /configure %cnfg%
1>nul 2>&1 del /q %cnfg%
if not exist "Tools\setup.exe" (
	1>nul 2>&1 copy %stpp% "Tools\setup.exe"
	1>nul 2>&1 del /q %stpp%
)
goto :eof

:usingPRSys

:: Global Variables
if /i '!xBit!' EQU '32' set "platform=86"
if /i '!xBit!' EQU '64' set "platform=64"
set "url=http://officecdn.microsoft.com/pr/!ProductId!"
set "identifier_2013=Retail_!langCd!_x-none"
set "identifier_2016=Volume.16_!langCd!_x-none"
set "identifier_c2r=!ProductYear!Volume.16_!langCd!_x-none"
set "identifier_c2r_R=!ProductYear!Retail.16_!langCd!_x-none"
set "exclude_2013=Retail.excludedapps"
set "exclude_2016=Retail.excludedapps.16"
set "exclude_c2r=!ProductYear!Volume.excludedapps.16"
set "exclude_c2r_R=!ProductYear!Retail.excludedapps.16=onedrive"
set "mShared=C:\Program Files\Common Files\Microsoft Shared"
set OfficeClickToRun="%mShared%\ClickToRun\OfficeClickToRun.exe"
if /i '!ProductYear!' EQU '2013' set OfficeClickToRun="%ProgramFiles%\Microsoft Office 15\ClientX!xBit!\OfficeClickToRun.exe"
set "misc_2013=updatesenabled=True autoUpgrade=True"
set "misc_2016=updatesenabled.16=True acceptalleulas.16=True displaylevel=True bitnessmigration=False"
set "unknown=flt.useoutlookshareaddon=unknown flt.useofficehelperaddon=unknown"

set source=Local
set baseurl="%fCD%"

if /i '!ProductYear!' EQU '2013' (
	if not exist !OfficeClickToRun! (
		1>nul 2>&1 rd /s /q "%ProgramFiles%\Microsoft Office 15\ClientX!xBit!"
		1>nul 2>&1 md "%ProgramFiles%\Microsoft Office 15\ClientX!xBit!"
		1>nul 2>&1 expand -f:*.* "Office\Data\!version!\i!xBit!!langID!.cab" "%ProgramFiles%\Microsoft Office 15\ClientX!xBit!"
	)
)

if /i '!ProductYear!' NEQ '2013' (
	if not exist !OfficeClickToRun! (
		1>nul 2>&1 rd /s /q "%mShared%\ClickToRun"
		1>nul 2>&1 md "%mShared%\ClickToRun"
		1>nul 2>&1 expand -f:*.* "Office\Data\!version!\i!xBit!0.cab" "%mShared%\ClickToRun"
	)
)

if not exist !OfficeClickToRun! ( echo. & echo Couldn't Extract OfficeClickToRun.exe File & echo. & goto :eof)

set "zeta="
if defined setupMode (
	echo !setupMode!|>nul find /i "2" && (
		for %%$ in (!selected_Products!) do (
			if not defined zeta (
				set zeta=true
				set "exclude_2013=!exclude_2013!=%%$"
				set "exclude_2016=!exclude_2016!=%%$"
				set "exclude_c2r=!exclude_c2r!=%%$"
			) else (
				set "exclude_2013=!exclude_2013!,%%$"
				set "exclude_2016=!exclude_2016!,%%$"
				set "exclude_c2r=!exclude_c2r!,%%$"
			)
		)
	)
)
if not defined zeta (
	set "exclude_2013=!exclude_2013!=onedrive"
	set "exclude_2016=!exclude_2016!=onedrive"
	set "exclude_c2r=!exclude_c2r!=onedrive"
)

if /i '!ProductYear!' EQU '2013' (
	set OkStatus=
	
	if defined setupMode (
		echo !setupMode!|>nul find /i "1" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=Excel%identifier_2013%^|PowerPoint%identifier_2013%^|Word%identifier_2013% productreleaseid=ExcelRetail,PowerPointRetail,WordRetail cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "2" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus%identifier_2013% productreleaseid=ProPlusRetail cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! ProPlus!exclude_2013! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "3" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=VisioPro%identifier_2013%^|ProjectPro%identifier_2013% productreleaseid=VisioProRetail,ProjectProRetail cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "5" && (
			set "productstoadd="
			set "productreleaseidNew="
			set "exclude_LIST="
			for %%$ in (!selected_Products!) do (
				if not defined productstoadd (
					set "productreleaseidNew=%%$Retail"
					set "productstoadd=%%$!identifier_2013!"
					set "exclude_LIST=%%$!exclude_2013!"
				) else (
					set "productreleaseidNew=!productreleaseidNew!,%%$Retail"
					set "productstoadd=!productstoadd!|%%$!identifier_2013!"
					set "exclude_LIST=!exclude_LIST! %%$!exclude_2013!"
				)
			)
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=!productstoadd! productreleaseid=!productreleaseidNew! cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "6" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus%identifier_2013%^|VisioPro%identifier_2013%^|ProjectPro%identifier_2013% productreleaseid=ProPlusRetail,VisioProRetail,ProjectProRetail cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! ProPlus!exclude_2013! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
			goto :eof
		)
	)
	
	(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus%identifier_2013% productreleaseid=ProPlusRetail cdnbaseurl=!url! baseurl=!baseurl! version=!version! mediatype=!source! scenario=unknown scenariosubtype=!source! lcid=!langID! !misc_2013!) && set OkStatus=true
	goto :eof
)

if /i '!ProductYear!' EQU '2016' (
	set OkStatus=
	
	if defined setupMode (
		echo !setupMode!|>nul find /i "1" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=Excel!identifier_2016!^|PowerPoint!identifier_2016!^|Word!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! Excel!exclude_2016! PowerPoint!exclude_2016! Word!exclude_2016! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "2" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! ProPlus!exclude_2016! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "3" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=VisioPro!identifier_2016!^|ProjectPro!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! VisioPro!exclude_2016! ProjectPro!exclude_2016! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "4" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=O365Business!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
			goto :eof
		)
		echo !setupMode!|>nul find /i "5" && (
			set "productstoadd="
			set "exclude_LIST="
			for %%$ in (!selected_Products!) do (
				if not defined productstoadd (
					set "exclude_LIST=%%$!exclude_2016!"
					set "productstoadd=%%$!identifier_2016!"
				) else (
					set "exclude_LIST=!exclude_LIST! %%$!exclude_2016!"
					set "productstoadd=!productstoadd!|%%$!identifier_2016!"
				)
			)
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=!productstoadd! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! !exclude_LIST! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true			
			goto :eof
		)
		
		echo !setupMode!|>nul find /i "6" && (
			(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_2016!^|VisioPro!identifier_2016!^|ProjectPro!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! ProPlus!exclude_2016! VisioPro!exclude_2016! ProjectPro!exclude_2016! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
			goto :eof
		)
	)
	
	(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_2016! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
	goto :eof
)

if defined setupMode (
	echo !setupMode!|>nul find /i "1" && (
		(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=Excel!identifier_c2r!^|PowerPoint!identifier_c2r!^|Word!identifier_c2r! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! Excel!exclude_c2r! PowerPoint!exclude_c2r! Word!exclude_c2r! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "2" && (
		(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_c2r! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! ProPlus!exclude_c2r! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "3" && (
		(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=VisioPro!identifier_c2r!^|ProjectPro!identifier_c2r! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! VisioPro!exclude_c2r! ProjectPro!exclude_c2r! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "5" && (
		set "productstoadd="
		for %%$ in (!selected_Products!) do (
			if not defined productstoadd (
				(echo "%%$"|>nul find /i "OneNote") && (
					set "productstoadd=%%$!identifier_c2r_R!"
					set "exclude_LIST=%%$!exclude_c2r_R!"
					) || (
					set "productstoadd=%%$!identifier_c2r!"
					set "exclude_LIST=%%$!exclude_c2r!"
				)
			) else (
				(echo "%%$"|>nul find /i "OneNote") && (
					set "productstoadd=!productstoadd!|%%$!identifier_c2r_R!"
					set "exclude_LIST=!exclude_LIST! %%$!exclude_c2r_R!"
					) || (
					set "productstoadd=!productstoadd!|%%$!identifier_c2r!"
					set "exclude_LIST=!exclude_LIST! %%$!exclude_c2r!"
				)
			)
		)
		(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=!productstoadd! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! !exclude_LIST! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
		goto :eof
	)
	
	echo !setupMode!|>nul find /i "6" && (
		(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_c2r!^|VisioPro!identifier_c2r!^|ProjectPro!identifier_c2r! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! ProPlus!exclude_c2r! VisioPro!exclude_c2r! ProjectPro!exclude_c2r! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
		goto :eof
	)
)

set OkStatus=
(!OfficeClickToRun! platform=x!platform! culture=!langCd! productstoadd=ProPlus!identifier_c2r! cdnbaseurl.16=!url! baseurl.16=!baseurl! version.16=!version! mediatype.16=!source! sourcetype.16=!source! !misc_2016! deliverymechanism=!ProductId! !unknown!) && set OkStatus=true
goto :eof

:UpdatePID
if /i '!ProductYear!' EQU '2013' (
	set "ProductId=39168D7E-077B-48E7-872C-B232C3E72675"
)

if /i '!ProductYear!' EQU '2016' (
	set "ProductId=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
)

if /i '!ProductYear!' EQU '2019' (
	set "ProductId=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
)

if /i '!ProductYear!' EQU '2021' (
	set "ProductId=5030841d-c919-4594-8d2d-84ae4f96e58e"
)
set VersionReleaseApi="https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/?audienceFFN=!ProductId!"
goto :eof

:GetFolderInfo
if exist Office\Data\v32.cab set lBit=32
if exist Office\Data\v64.cab set lBit=64
if exist Office\Data\v32.cab (if exist Office\Data\v64.cab set lBit=!sBit!)

set "multiLang="
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
	if exist "office\data\!version!\i%lBit%%%h.cab" (
		set "lLngID=%%h"
		if not defined multiLang (
			set "multiLang=!lLngID!"
		) else (
			(echo '!multiLang!' |>nul find /i "!lLngID!") || (set "multiLang=!multiLang!,!lLngID!")
		)
	)
)

if defined lLngID (
	if defined multiLang (
	
		set "count=0"
		set "countVal="
		if /i '!lLngID!' NEQ '!multiLang!' (
			echo:
			echo Multi Language Found.
			echo _____________________
			echo:
			for %%# in (!multiLang!) do (
				set /a count+=1
				call :FindLngId %%#
				echo Language [!count!] :: [%%#] !langIdName!
				set "lang_!count!=%%#"
				set "countVal=!countVal!!count!"
			)
			
			echo:
			CHOICE /C !countVal! /M "Select Language ID ::"
			
			FOR /L %%# IN (1,1,!count!) DO (			
				if /i '!errorlevel!' EQU '%%#' (
					set "lLngID=!lang_%%#!"
				)
			)
		)
	)
)

REM for %%g in (2013, 2016, 2019, 2021) do (
	REM ( echo %fCD% | find /i "%%g">nul && (set "lYear=%%g") ) || ( vol | find /i "%%g">nul && (set "lYear=%%g") )
REM )

set "Ignore="
>nul 2>&1 dir /ad /b "office\data\16*" && (
	for /f "tokens=*" %%g in ('dir /ad /b "office\data\16*"') do set "version=%%g"
)

echo "!version!" | >nul findstr /r "1[5-6].[0-9].[0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9][0-9]" || (
	echo "!version!" | >nul findstr /r "1[5-6].[0-9].[0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9]" || 	(
		goto :eof
	)
)

echo !xBit! | >nul find /i "32" && set "CabFile=Office\Data\v32.cab"
echo !xBit! | >nul find /i "64" && set "CabFile=Office\Data\v64.cab"

(>nul expand !CabFile! -F:VersionDescriptor.xml "%temp%")  || goto :eof
if not exist "%temp%\VersionDescriptor.xml" goto :eof

set "DeliveryMechanism="
for /f "tokens=*" %%# in ('type "%temp%\VersionDescriptor.xml"^|find /i "DeliveryMechanism"') do set "DeliveryMechanism=%%#"

if not defined DeliveryMechanism (
	type "%temp%\VersionDescriptor.xml"| >nul find /i "15." && (
		set "lYear=2013"
		set "ProductYear=2013"
		set "ProductId=39168D7E-077B-48E7-872C-B232C3E72675"
	)
	goto :eof
)
set "DeliveryMechanism=!DeliveryMechanism:~28,-2!"

rem %%g Name, %%h Channel
>"%temp%\tmp" call :Channel_List
for /f "tokens=1,2 delims=*" %%g in ('type "%temp%\tmp"') do (
	(echo !DeliveryMechanism! | >nul find /i "%%h") && (
		set "ProductId=%%h"
		set "lYear=2016"
		set "ProductYear=2016"
		
		echo "%%g" | >nul find /i "PerpetualVL2019" && set "ProductYear=2019" & set "lYear=2019"
		echo "%%g" | >nul find /i "PerpetualVL2021" && set "ProductYear=2021" & set "lYear=2021"
		set "Ignore=*"
	)
)
goto :eof

:UpdateFolderInfo
echo.
if not defined lBit ( echo Error Pharse Disk System Type. use default.) else (echo Disk System Type = x%lBit%)
if not defined lYear ( echo Error Pharse Disk Product Year. use default.) else (echo Disk Product Year = %lYear%)
if not defined lLngID ( echo Error Pharse Disk Language Id. use default.) else (echo Disk Language Id = %lLngID%)

if defined lBit		call :updateBit %lBit%
if defined lYear	call :updateYear %lYear%
if defined lLngID	call :convertLngId %lLngID%
goto :eof

:convertLngId
set var=&set var=%*
if defined var (

	:: %%g=English %%h=1033 %%i=en-us %%j:0409
	>"%temp%\tmp" call :Language_List
	for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
		if /i '%var%' EQU '%%h' (
			call :UpdateLanguage %%h %%i
			goto :convertLngId_
		)
	)
)
:convertLngId_
goto :eof

:FindLngId
set "langIdName="
set var=&set var=%*
if defined var (

	:: %%g=English %%h=1033 %%i=en-us %%j:0409
	>"%temp%\tmp" call :Language_List
	for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
		if /i '%var%' EQU '%%h' (
			set "langIdName=%%g"
			goto :FindLngId_
		)
	)
)
:FindLngId_
goto :eof

:: PS Helpers
:: PS Helpers
:: PS Helpers

:Query
set args_1=%1
set args_2=%2
>nul 2>&1 del /q "%temp%\result"
>"%temp%\script.ps1" echo Get-WmiObject -Query "!args_1:~1,-1!" ^| format-table -Property !args_2:~1,-1!
>"%temp%\result" 2>&1 powershell -noprofile -executionpolicy bypass -file "%temp%\script.ps1"
goto :eof

:Query_Invoke_NoArgs
set args_1=%1
set args_2=%2
>nul 2>&1 del /q "%temp%\result"
>"%temp%\script.ps1" echo Get-WmiObject -Query "!args_1:~1,-1!" ^| ForEach-Object { ($_).!args_2:~1,-1!()}
>"%temp%\result" powershell -noprofile -executionpolicy bypass -file "%temp%\script.ps1"
goto :eof

:Query_Invoke_args
set args_1=%1
set args_2=%2
set args_3=%3
>nul 2>&1 del /q "%temp%\result"
>"%temp%\script.ps1" echo Get-WmiObject -Query "!args_1:~1,-1!" ^| ForEach-Object { ($_).!args_2:~1,-1!("!args_3:~1,-1!")}
>"%temp%\result" powershell -noprofile -executionpolicy bypass -file "%temp%\script.ps1"
goto :eof

:Channel_List
echo Current*492350f6-3a01-4f97-b9c0-c7c6ddf67d60
echo CurrentPreview*64256afe-f5d9-4f86-8936-8840a6a4f5be
echo BetaChannel*5440fd1f-7ecb-4221-8110-145efaa6372f
echo MonthlyEnterprise*55336b82-a18d-4dd6-b5f6-9e5095c314a6
echo SemiAnnual*7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
echo SemiAnnualPreview*b8f9b850-328d-4355-9145-c59439a0c4cf
echo DogfoodDevMain*ea4a4090-de26-49d7-93c1-91bff9e53fc3
echo Manual_Override*ea4a4090-de26-49d7-93c1-91bff9e53fc3
echo Manual_Override*f3260cf1-a92c-4c75-b02e-d64c0a86a968
echo Manual_Override*c4a7726f-06ea-48e2-a13a-9d78849eb706
echo Manual_Override*834504cc-dc55-4c6d-9e71-e024d0253f6d
echo Manual_Override*5462eee5-1e97-495b-9370-853cd873bb07
echo Manual_Override*f4f024c8-d611-4748-a7e0-02b6e754c0fe
echo Manual_Override*b61285dd-d9f7-41f2-9757-8f61cba4e9c8
echo Manual_Override*9a3b7ff2-58ed-40fd-add5-1e5158059d1c
echo Manual_Override*86752282-5841-4120-ac80-db03ae6b5fdb
echo Manual_Override*2e148de9-61c8-4051-b103-4af54baffbb4
echo Manual_Override*12f4f6ad-fdea-4d2a-a90f-17496cc19a48
echo Manual_Override*0002c1ba-b76b-4af9-b1ee-ae2ad587371f
echo PerpetualVL2019*f2e724c1-748f-4b47-8fb8-8e0d210e9208
echo PerpetualVL2021*5030841d-c919-4594-8d2d-84ae4f96e58e
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