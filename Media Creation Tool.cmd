@cls
@echo off
>nul chcp 437

mode 130, 30
set cnfg="%temp%\tmp.xml"
set stpp="Tools\setup.exe"
REM set Windows_7_Support=true
set user_agent=Microsoft-Delivery-Optimization/10.0
set officecdn=officecdn.microsoft.com.edgesuite.net/pr
SETLOCAL EnableDelayedExpansion
title Media Creation Tool

echo "%~dp0"|>nul findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
    Echo Invalid path: "%~dp0"
    Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	pause
	exit /b
)

:: basic validtion
call :cleanVariable

:::: Run as Admin with native shell, any path, params, loop guard, minimal i/o, by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\" &call \"%%2\" %%3" & set "_= %*"
>nul fltmc || if "%f0%" neq "%~f0" ( cd.>"%tmp%\runas.Admin" & start "%~n0" /high "%tmp%\runas.Admin" "%~f0" "%_:"=""%" &exit /b )

cd /d "%~dp0"

call :cleanOffice
call :cleanFiles

set xTmp=
set xTmp=%*
if not defined xTmp (
	if not exist BuildInfo.ini (
		echo.
		echo Tool to create ISO file,
		echo of your latest chosen office version,
		echo on your Desktop.
		echo.
		echo How to Use ^?
		echo Extract It to folder Or Mount It.
		echo [If you extract It, keep Origional ISO Name]
		echo And Use "Local tool" to Install it
		echo.
		echo Supprted C2R Versions 	:: [ 2013, 2016, 2019, 2021 ] [ Fallback -- 2021 ]
		echo Supprted System Bit 	:: [ 32, 64, Multi ] [ Fallback -- System Type ]
		echo Supprted Language List  :: [ Open "Read Me" File ] [ Fallback -- System Language ]
		echo.
		echo Optional Parameters [Year] [Bit] [Language]
		echo Can Be Nothing Or Everything, up to you
		echo.
		echo "%~n0" 			  --   Office 2021 -- [ System Type ] -- [ System Language ]
		echo "%~n0" 2016 64 		  --   Office 2016 -- X64 -- [ System Language ]
		echo "%~n0" 2013 multi hebrew   --   Office 2013 -- X32-X64 -- Hebrew
		echo.
		CHOICE /C CE /M "Press C For Continue, E for Exit."
		if !errorlevel! EQU 2 exit /b
		cls
	)
)

echo.
echo Check parameters
echo:

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
rem St[3] language - user prompt
set "xVal="&set "xVal=%*"
if defined xVal call :InputPharser !xVal!
rem St[4] read from config file
call :UpdateBuildInfo
rem done checking

if not defined userSelected (
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
	echo:
)

:: 2 option
REM goto :usingODT
goto :usingPRSys
:next_

if defined DownloadFailed (
	echo.
	echo Downlaod Failed ...
	call :cleanOffice
	goto :end
)

call :updateIsoFile
call :BuildIsoFile !sBit!
call :cleanOffice

if not exist "%USERPROFILE%\desktop\%iso_file%" (
	echo.
	echo Fail to create Office Iso File
	goto :end
)

echo Create "%iso_file%" File Succeed
echo Have a Great Day ...
goto :eof

:end
echo.
echo please wait 5 seconds..
SETLOCAL DisableDelayedExpansion
timeout /t 6 /NOBREAK
goto :eof

:cleanVariable
set vars=DownloadFailed, xVal, xBit, sBit, sLng, lLngID, lBit, lLng, lYear, equal, multi, stream, langId, LangName, langCd, version, cscript, platform, file_list, iso_file, iso_label ,ProductId, SysLanCD, SysLanIdHex, ProductYear, SupportedBit, VersionReleaseApi, SupportedbuildYear
for %%v in (%vars%) do (set %%v=)
goto :eof

:stpDownload
echo.&echo Downloading Setup file
1>nul 2>&1 tools\wget --quiet --tries=3 --retry-connrefused --output-document=%stpp% "https://officecdn.microsoft.com/pr/wsus/setup.exe"
goto :eof

:stpDownloadv2
echo Download File :: setup.exe
1>nul 2>&1 tools\wget --quiet --tries=3 --retry-connrefused --output-document=%stpp% "https://officecdn.microsoft.com/pr/wsus/setup.exe"
goto :eof

:UpdateBuildInfo
set "SupportedBit=32, 64, Multi"
set "SupportedbuildYear=2013, 2016, 2019, 2021"
if not exist BuildInfo.ini goto :eof
for /f "tokens=1,2* delims==" %%g in (BuildInfo.ini) do call :BuildInfoPharser %%g %%h

rem Just in case of ..
call :LangPharser !LangName!

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
set ProductYear=%*
goto :eof

:updateBit
set xBit=%*
if /i '%*' EQU 'Multi' (
	set multi=true
	set xBit=32
)
goto :eof

:ZeroValues
call :updateBit %1
call :updateYear %2
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
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
	call :compareXY "%*" "%%g"
	if defined equal (
		set "LangName=%%g"
		call :UpdateLanguage %%h %%i
		call :UpdateFileList
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
		call :UpdateFileList
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
set "UpdateLangCheck="
if defined multi (
	call :compareXY "%pharse:~11%" ""
	if not defined equal (
		call :LangPharser %pharse:~11%
		set UpdateLangCheck=***
	)
)

if not defined multi (
	call :compareXY "%pharse:~8%" ""
	if not defined equal (
		call :LangPharser %pharse:~8%
		set UpdateLangCheck=***
	)
)

if not defined UpdateLangCheck (
	call :LangPharser !LangName!
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

:: Actual diffrents from scripts Moved here
:: its more elegent way 

:BuildConfigartionFile
if /i '%productyear%' EQU '2016' (
	echo Build Configuration File for Office %ProductYear% x%1 !langCd!
	>%cnfg%  echo ^<Configuration^>
	if defined Windows_7_Support 	 	>>%cnfg% echo ^<Add OfficeClientEdition^=^"%1^" Channel^=^"Current^" Version^=^"16.0.12527.22021^"^>
	if not defined Windows_7_Support 	>>%cnfg% echo ^<Add OfficeClientEdition^=^"%1^" Channel^=^"Current^"^>
	>>%cnfg% echo     ^<Product ID^=^"ProPlusRetail^"^>
	>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
	>>%cnfg% echo     ^<^/Product^>
	>>%cnfg% echo ^<^/Add^>
	>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
	>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL%ProductYear%^" ^/^>
	>>%cnfg% echo ^<^/Configuration^>
) else (
	echo Build Configuration File for Office %ProductYear% x%1 !langCd!
	>%cnfg%  echo ^<Configuration^>
	>>%cnfg% echo ^<Add OfficeClientEdition^=^"%1^" Channel^=^"PerpetualVL%ProductYear%^"^>
	>>%cnfg% echo     ^<Product ID^=^"ProPlus%ProductYear%Volume^"^>
	>>%cnfg% echo       ^<Language ID^=^"!langCd!^" ^/^>
	>>%cnfg% echo     ^<^/Product^>
	>>%cnfg% echo ^<^/Add^>
	>>%cnfg% echo ^<Display Level^=^"Full^" AcceptEULA^=^"TRUE^" ^/^>
	>>%cnfg% echo ^<Updates Enabled^=^"TRUE^" Channel^=^"PerpetualVL%ProductYear%^" ^/^>
	>>%cnfg% echo ^<^/Configuration^>
)
goto :eof

:usingODT
if not exist %stpp% call :stpDownload

if defined multi (
	set msg=
	if /i '%productyear%' EQU '2013' set msg=Download office %ProductYear% vl %langCd:~0,2% x32-x64 v%version%
	if /i '%productyear%' EQU '2016' set msg=Download office %ProductYear% vl %langCd:~0,2% x32-x64 v%version%
	if not defined msg set msg=Download office %ProductYear% Perpetual vl %langCd:~0,2% x32-x64 v%version%
	echo:
	echo !msg!
	
	call :officeDownload 32
	call :officeDownload 64
)
if not defined multi (
	set msg=
	if /i '%productyear%' EQU '2013' set msg=Download office %ProductYear% vl %langCd:~0,2% x!xBit! v%version%
	if /i '%productyear%' EQU '2016' set msg=Download office %ProductYear% vl %langCd:~0,2% x!xBit! v%version%
	if not defined msg set msg=Download office %ProductYear% Perpetual vl %langCd:~0,2% x!xBit! v%version%
	echo:
	echo !msg!
	
	call :officeDownload !xBit!
)

if not exist "office" (
	echo.
	echo Not found any Local Installation
	goto :end
)

if not exist "office\data" (
	echo.
	echo !~~ Folder Error ~~!
	call :cleanOffice
	goto :end
)

for /f %%g in ('dir /ad/b office\data') do set version=%%g
if not defined version (
	echo.
	echo Could not found office version
	call :cleanOffice
	goto :end
)
goto :next_

:usingPRSys
call :UpdatePID

:: Find Version Function ~ Start
echo Verify Version

1>nul 2>&1 del /q /s "%temp%\v!xBit!.cab"
1>nul 2>&1 del /q /s "%temp%\VersionDescriptor.xml"

:: Nethood 1XX
if /i '!productyear!' NEQ '2013' (
	1>nul 2>&1 tools\wget --quiet --tries=3 --retry-connrefused --user-agent=%user_agent% --output-document="%temp%\VersionDescriptor.xml" %VersionReleaseApi%
)

:: Nethood 2XX
if /i '!productyear!' EQU '2013' (
	1>nul 2>&1 tools\wget --quiet --user-agent=%user_agent% --output-document="%temp%\v!xBit!.cab" "officecdn.microsoft.com.edgesuite.net/pr/!ProductId!/Office/Data/v!xBit!.cab"
	1>nul 2>&1 expand -f:VersionDescriptor* "%temp%\v!xBit!.cab" "%temp%"
)

if not exist "%temp%\VersionDescriptor.xml" (
	echo.
	echo Check Your internet connection.
	goto :end
)

:: Nethood 1XX
if /i '!productyear!' NEQ '2013' (
	for /f "tokens=2 delims=:, " %%g in ('type "%temp%\VersionDescriptor.xml" ^| find /i "AvailableBuild"') do (
		if not defined version (
			set "arg_=%%g"
			set "version=!arg_:~1,-1!"
			set /p version=Set Office Build - or press return for !version! :: 
			REM echo:
		)	
	)
)

:: Nethood 2XX
if /i '!productyear!' EQU '2013' (
	for /f  "tokens=2 delims= " %%g in ('type "%temp%\VersionDescriptor.xml" ^| find /i "Available Build"') do (
		if not defined version (
			set "version=%%g"
			set "version=!version:~7,-1!"
			set /p version=Set Office Build - or press return for !version! :: 
			REM echo:
		)
	)
)

1>nul 2>&1 del /q /s "%temp%\v!xBit!.cab"
1>nul 2>&1 del /q /s "%temp%\VersionDescriptor.xml"

if not defined version (
	echo.
	echo Could not found office version
	goto :end
)

if /i '%ProductYear%' EQU '2016' (
	if defined Windows_7_Support (
		set version=16.0.12527.22021
	)
)

if defined multi (
	set msg=
	if /i '%productyear%' EQU '2013' set msg=Download office %ProductYear% vl %langCd:~0,2% x32-x64 v%version%
	if /i '%productyear%' EQU '2016' set msg=Download office %ProductYear% vl %langCd:~0,2% x32-x64 v%version%
	if not defined msg set msg=Download office %ProductYear% Perpetual vl %langCd:~0,2% x32-x64 v%version%
	echo:
	echo !msg!
)
if not defined multi (
	set msg=
	if /i '%productyear%' EQU '2013' set msg=Download office %ProductYear% vl %langCd:~0,2% x!xBit! v%version%
	if /i '%productyear%' EQU '2016' set msg=Download office %ProductYear% vl %langCd:~0,2% x!xBit! v%version%
	if not defined msg set msg=Download office %ProductYear% Perpetual vl %langCd:~0,2% x!xBit! v%version%	
	echo:
	echo !msg!
)

md "Office\Data\%version%"
call :cleanFiles

:: actual download
if not exist %stpp% call :stpDownloadv2
call :DownloadBaseFile
for %%i in (%file_list%) do (
	if not defined DownloadFailed (
		call :DownloadModule %%i
	)
)

goto :next_

:UpdatePID
if /i '%ProductYear%' EQU '2013' (
	set ProductId=39168D7E-077B-48E7-872C-B232C3E72675
)
if /i '%ProductYear%' EQU '2016' (
	set ProductId=492350f6-3a01-4f97-b9c0-c7c6ddf67d60
)
if /i '%ProductYear%' EQU '2019' (
	set ProductId=f2e724c1-748f-4b47-8fb8-8e0d210e9208
)
if /i '%ProductYear%' EQU '2021' (
	set ProductId=5030841d-c919-4594-8d2d-84ae4f96e58e
)
set VersionReleaseApi="https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/?audienceFFN=!ProductId!"
goto :eof

:DownloadBaseFile
echo.
echo LOCAL  Path :: %cd%
echo CDN    Path :: %officecdn%/!ProductId!

if defined multi (
	call :DownloadBaseFileHelper 32
	call :DownloadBaseFileHelper 64
)
if not defined multi (
	call :DownloadBaseFileHelper !xBit!
)
goto :eof

:DownloadBaseFileHelper
echo.
echo CDN/Office/Data/v%1.cab --^> LOCAL\Office\Data\v%1.cab
REM tools\curl -# -o "Office\Data\v%1.cab" "%officecdn%/!ProductId!/Office/Data/v%1.cab"
tools\wget --quiet --show-progress --user-agent=%user_agent% --output-document="Office\Data\v%1.cab" "%officecdn%/!ProductId!/Office/Data/v%1.cab"
if %errorlevel% NEQ 0 (
	1>nul 2>&1 del /q "Office\Data\v%1.cab"
	set DownloadFailed=true
	goto :eof
)

echo.
echo CDN/Office/Data/v%1_%version%.cab --^> LOCAL\Office\Data\v%1_%version%.cab
REM tools\curl -# -o "Office\Data\v%1_%version%.cab" "%officecdn%/!ProductId!/Office/Data/v%1_%version%.cab"
tools\wget --quiet --show-progress --user-agent=%user_agent% --output-document="Office\Data\v%1_%version%.cab" "%officecdn%/!ProductId!/Office/Data/v%1_%version%.cab"
if %errorlevel% NEQ 0 (
	1>nul 2>&1 del /q "Office\Data\v%1_%version%.cab"
	set DownloadFailed=true
	goto :eof
)

goto :eof

:DownloadModule
echo.
echo CDN/Office/Data/%version%/%* --^> LOCAL\Office\Data\%version%\%*
REM tools\curl -# -o "Office\Data\%version%\%*" "%officecdn%/!ProductId!/Office/Data/%version%/%*"
tools\wget --quiet --show-progress --user-agent=%user_agent% --output-document="Office\Data\%version%\%*" "%officecdn%/!ProductId!/Office/Data/%version%/%*"
if %errorlevel% NEQ 0 (
	1>nul 2>&1 del /q "Office\Data\%version%\%*"
	if /i '%productyear%' NEQ '2021' (echo %*|>nul find /i ".cat" && goto :eof)
	if /i '%productyear%' EQU '2013' (echo %*|>nul find /i "i320.cab" && goto :eof)
	if /i '%productyear%' EQU '2013' (echo %*|>nul find /i "i640.cab" && goto :eof)
	if /i '%productyear%' EQU '2013' (echo %*|>nul find /i "i640.cab" && goto :eof)
	if /i '%productyear%' EQU '2013' (echo %*|>nul find /i "sp32!langID!.cab" && goto :eof)
	if /i '%productyear%' EQU '2013' (echo %*|>nul find /i "sp64!langID!.cab" && goto :eof)
	set DownloadFailed=true
	goto :eof
)
goto :eof

:UpdateFileList
REM echo.&echo Enter UpdateFileList
if defined multi (
	REM echo Enter Multi
	call :CleanFileList
	call :UpdateFileListHelper 32
	call :UpdateFileListHelper 64
	goto :eof
)
if not defined multi (
	REM echo Enter Not Multi
	call :CleanFileList
	call :UpdateFileListHelper !xBit! Addins
	goto :eof
)
goto :eof

:CleanFileList
set file_list=
goto :eof

:UpdateFileListHelper
call :updateStreamChannel %1
if not defined file_list (
	REM echo Create File List - x%1 - %stream%
	set file_list=i%10.cab,i%10.cab.cat,i%1!langID!.cab,s%10.cab,s%1!langID!.cab,%stream%.!langCd!.dat,%stream%.!langCd!.dat.cat,%stream%.x-none.dat,%stream%.x-none.dat.cat,sp%1!langID!.cab
	goto :UpdateFileListHelper_Next
)

if defined file_list (
	REM echo Update File List - x%1 - %stream%
	set file_list=%file_list%,i%10.cab,i%10.cab.cat,i%1!langID!.cab,s%10.cab,s%1!langID!.cab,%stream%.!langCd!.dat,%stream%.!langCd!.dat.cat,%stream%.x-none.dat,%stream%.x-none.dat.cat,sp%1!langID!.cab
	goto :UpdateFileListHelper_Next
)

:UpdateFileListHelper_Next
if /i '%2' EQU 'Addins' (
	if /i '%1' EQU '32' (
		:: even 32Bit release needed some x64 files
		REM echo echo Update File List - Missing x64 Files
		set file_list=%file_list%,i64!langID!.cab,i640.cab,i640.cab.cat
	)
)
goto :eof

:updateStreamChannel
if /i '%1' EQU '32' set stream=stream.x86
if /i '%1' EQU '64' set stream=stream.x64
goto :eof


:officeDownload
1>nul 2>&1 del /q %cnfg%
call :BuildConfigartionFile %1
echo download Files From Ms Servers using Odt Tool
if not exist %cnfg% goto :eof
%stpp% /download %cnfg%
goto :eof

:cleanFiles
1>nul 2>&1 del /q tmp.xml
1>nul 2>&1 del /q "%temp%\v!xBit!.cab"
1>nul 2>&1 del /q "%temp%\VersionDescriptor.xml"
goto :eof

:cleanOffice
1>nul 2>&1 rd /s/q "office"
goto :eof

:updateIsoFile
if defined multi 		set iso_file=%langCd:~0,2%_office_%ProductYear%_C2R_vl_x32-x64_v%version%.iso
if not defined multi 	set iso_file=%langCd:~0,2%_office_%ProductYear%_C2R_vl_x!xBit!_v%version%.iso

if defined multi 		set iso_label=%langCd:~0,2%_office_%ProductYear%_C2R_vl_x32-x64
if not defined multi 	set iso_label=%langCd:~0,2%_office_%ProductYear%_C2R_vl_x!xBit!

goto :eof

:BuildIsoFile
if /i '%1' NEQ '' (
	echo.
	echo Build Iso File ....
	1>nul 2>&1 del /q "%USERPROFILE%\desktop\%iso_file%"
	if exist "%USERPROFILE%\desktop\%iso_file%" (
		echo.
		echo ERROR :: file in use
		echo "%USERPROFILE%\desktop\%iso_file%"
		goto :eof
	)
	
	1>nul 2>&1 pushd "%windir%\System32\WindowsPowershell\v1.0\" && (
		
		popd
		
		rem lean and mean snippet by AveYo, 2021 
		rem export directory as (bootable) udf iso
		1>nul call :DIR2ISO "%cd%" "%USERPROFILE%\desktop\%iso_file%" !iso_label!
		
	) || (
	
		rem MS oscdimg Tool
		1>nul 2>&1 x%1\oscdimg.exe -m -u1 -L%iso_label% "%cd%" "%USERPROFILE%\desktop\%iso_file%"
	)

	goto :eof
)

if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 	call :BuildIsoFile 32)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 		call :BuildIsoFile 64)
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	call :BuildIsoFile 64
if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	call :BuildIsoFile 64
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

rem lean and mean snippet by AveYo, 2021 
rem export directory as (bootable) udf iso
$:DIR2ISO: #:: [PARAMS] "directory" "file.iso" ["label"]
set ^ #="$f0=[io.file]::ReadAllText($env:0);$0=($f0-split'\$%0:.*')[1];$1=$env:1-replace'([`@$])','`$1';iex(\"$0 `r`n %0 $1\")"
set ^ #=& set "0=%~f0"& set 1=%*& powershell -nop -c %#%& exit/b %errorcode%
function :DIR2ISO ($dir, $iso, $vol='DVD_ROM') {if (!(test-path -Path $dir -pathtype Container)) {"[ERR] $dir\";exit 1}; $code=@"
 using System; using System.IO; using System.Runtime.Interop`Services; using System.Runtime.Interop`Services.ComTypes;
 public class dir2iso {public int AveYo=2021; [Dll`Import("shlwapi",CharSet=CharSet.Unicode,PreserveSig=false)]
 internal static extern void SHCreateStreamOnFileEx(string f,uint m,uint d,bool b,IStream r,out IStream s);
 public static int Create(string file, ref object obj, int bs, int tb) { IStream dir=(IStream)obj, iso;
 try {SHCreateStreamOnFileEx(file,0x1001,0x80,true,null,out iso);} catch(Exception e) {Console.WriteLine(e.Message); return 1;}
 int d=tb>1024 ? 1024 : 1, pad=tb%d, block=bs*d, total=(tb-pad)/d, c=total>100 ? total/100 : total, i=1, MB=(bs/1024)*tb/1024;
 Console.Write("{0,2}%  {1}MB {2}  :DIR2ISO",0,MB,file); if (pad > 0) dir.CopyTo(iso, pad * block, Int`Ptr.Zero, Int`Ptr.Zero);
 while (total-- > 0) {dir.CopyTo(iso, block, Int`Ptr.Zero, Int`Ptr.Zero); if (total % c == 0) {Console.Write("\r{0,2}%",i++);}}
 iso.Commit(0); Console.WriteLine("\r{0,2}%  {1}MB {2}  :DIR2ISO", 100, MB, file); return 0;} }
"@; & { $cs=new-object CodeDom.Compiler.CompilerParameters; $cs.GenerateInMemory=1 #:: no`warnings
 $compile=(new-object Microsoft.CSharp.CSharpCodeProvider).CompileAssemblyFromSource($cs, $code)
 $BOOT=@(); $bootable=0; $mbr_efi=@(0,0xEF); $images=@('boot\etfsboot.com','efi\microsoft\boot\efisys.bin') #:: efisys_noprompt
 0,1|% { $bootimage=join-path $dir -child $images[$_]; if (test-path -Path $bootimage -pathtype Leaf) {
 $bin=new-object -ComObject ADODB.Stream; $bin.Open(); $bin.Type=1; $bin.LoadFromFile($bootimage)
 $opt=new-object -ComObject IMAPI2FS.BootOptions; $opt.AssignBootImage($bin.psobject.BaseObject); $opt.Manufacturer='Microsoft'
 $opt.PlatformId=$mbr_efi[$_]; $opt.Emulation=0; $bootable=1; $BOOT += $opt.psobject.BaseObject } }
 $fsi=new-object -ComObject IMAPI2FS.MsftFileSystemImage; $fsi.FileSystemsToCreate=4; $fsi.FreeMediaBlocks=0
 if ($bootable) {$fsi.BootImageOptionsArray=$BOOT}; $CONTENT=$fsi.Root; $CONTENT.AddTree($dir,$false); $fsi.VolumeName=$vol
 $obj=$fsi.CreateResultImage(); $r=[dir2iso]::Create($iso,[ref]$obj.ImageStream,$obj.BlockSize,$obj.TotalBlocks) };[GC]::Collect()
} $:DIR2ISO: #:: export directory as (bootable) udf iso - lean and mean snippet by AveYo, 2021