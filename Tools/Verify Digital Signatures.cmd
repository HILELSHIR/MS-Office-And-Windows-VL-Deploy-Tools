@(echo off% <#%) &color 07 &title Verify Digital Signatures - Files/Folders SendTo menu entry by AveYo
set "0=%~f0" &set 1=%*& powershell -nop -c iex ([io.file]::ReadAllText($env:0)) && exit /b ||#>)[1]

#,# Process command line arguments - supports multiple files and folders
  $arg = ([regex]'"[^"]+"|[^ ]+').Matches($env:1)
  $val = Get-Item -force -lit ($arg[0].Value.Trim('"'))
  $dir = Split-Path $val; cd -lit $dir
#,# Grab target files names
  $files = @()
  foreach ($a in $arg) {
    $f = gi -force -lit $a.Value.Trim('"')
    if ($f.PSTypeNames -match 'FileInfo') { $files += $f }
    else { dir -lit $f -rec -force |? { !$_.PSIsContainer } |% { $files += $_ } }
  }
#,# Verify digital signatures via built-in powershell
$err = @(); $filter = '.exe','.dll','.ocx','.sys','.cpl','.scr','.msi','.Msix','.appx','.appxbundle','.cab'
$files | foreach-object {
  if ($filter -contains $_.Extension) {
    $sig = Get-AuthenticodeSignature $_.FullName
    if ($sig.status -eq 0) { write-output $sig ($sig.SignerCertificate.Subject-split'"')[1] '' }
    else { $err += $_ | select-object FullName,LastWriteTime,Length,@{Name='Status';Expression={("Not signed / Invalid")}} }
  }
}
write-host
if ($err.length -eq 0) { write-host -fore yellow -back darkgreen " OK! " }
else { write-output $err; write-host -fore yellow -back darkred " ERR! " }
#,# Done

