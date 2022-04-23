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
echo -- Integrate Mondo 2016 Retail License
"!root!\Integration\integrator" /I /License PRIDName=MondoRetail.16 PidKey=2N6B3-BXW6B-W2XBT-VVQ64-7H7DH

echo -- Clean Registry Keys
for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (
    1>nul 2>&1 reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Microsoft\Office" /f
    1>nul 2>&1 reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Wow6432Node\Microsoft\Office" /f
    1>nul 2>&1 reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides" /f
)

echo -- Install Visual UI Registry Keys
call :reg_own "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" "" S-1-5-32-544 "" Allow SetValue
for %%# in (Word, Excel, Powerpoint, Access, Outlook, Publisher, OneNote, project, visio) do (
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true"
  1>nul 2>&1 reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" /f /v "%%#" /t REG_SZ /d "{FBDB3E18-A8EF-4FB3-9183-DFFD60BD0984},{CE5FFCAF-75DA-4362-A9CB-00D2689918AA},"
)
call :reg_own "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" "" S-1-5-32-544 "" Deny SetValue

echo -- Done.
echo.

echo Note:
echo To initiate the Visual Refresh,
echo it may be required to start some Office apps
echo a couple of times.
echo.
echo Many Thanks to Xtreme21, Krakatoa, rioachim
echo for helping make and debug this script
echo.

pause
exit /b

:reg_own #key [optional] all user owner access permission  :        call :reg_own "HKCU\My" "" S-1-5-32-544 "" Allow FullControl
powershell -nop -c $A='%~1','%~2','%~3','%~4','%~5','%~6';iex(([io.file]::ReadAllText('%~f0')-split':Own1\:.*')[1])&exit/b:Own1:
$D1=[uri].module.gettype('System.Diagnostics.Process')."GetM`ethods"(42) |where {$_.Name -eq 'SetPrivilege'} #`:no-ev-warn
'SeSecurityPrivilege','SeTakeOwnershipPrivilege','SeBackupPrivilege','SeRestorePrivilege'|foreach {$D1.Invoke($null, @("$_",2))}
$path=$A[0]; $rk=$path-split'\\',2; $HK=gi -lit Registry::$($rk[0]) -fo; $s=$A[1]; $sps=[Security.Principal.SecurityIdentifier]
$u=($A[2],'S-1-5-32-544')[!$A[2]];$o=($A[3],$u)[!$A[3]];$w=$u,$o |% {new-object $sps($_)}; $old=!$A[3];$own=!$old; $y=$s-eq'all'
$rar=new-object Security.AccessControl.RegistryAccessRule( $w[0], ($A[5],'FullControl')[!$A[5]], 1, 0, ($A[4],'Allow')[!$A[4]] )
$x=$s-eq'none';function Own1($k){$t=$HK.OpenSubKey($k,2,'TakeOwnership');if($t){0,4|%{try{$o=$t.GetAccessControl($_)}catch{$old=0}
};if($old){$own=1;$w[1]=$o.GetOwner($sps)};$o.SetOwner($w[0]);$t.SetAccessControl($o); $c=$HK.OpenSubKey($k,2,'ChangePermissions')
$p=$c.GetAccessControl(2);if($y){$p.SetAccessRuleProtection(1,1)};$p.ResetAccessRule($rar);if($x){$p.RemoveAccessRuleAll($rar)}
$c.SetAccessControl($p);if($own){$o.SetOwner($w[1]);$t.SetAccessControl($o)};if($s){$subkeys=$HK.OpenSubKey($k).GetSubKeyNames()
foreach($n in $subkeys){Own1 "$k\$n"}}}};Own1 $rk[1];if($env:VO){get-acl Registry::$path|fl} #:Own1: lean & mean snippet by AveYo