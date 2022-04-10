<#
.SYNOPSIS
    Wrapper specific for computer in workshop used as backup destination.

    Download latest version from Github and run as administrator with predefined variables
.EXAMPLE
    . sync-ToWorkshow.ps1

#>
param (
    $Path='\\pihl-fs\Pihl\',
    $DestinationPath='D:\Backup\Documents\',
    $DestinationExtraFiles='D:\BackupPihl-ExtraFiles\',
    $uriBackup = 'https://raw.githubusercontent.com/KlasPihl/sync-directory/main/sync-directory.ps1'
)
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
  $arguments = "& '" +$myinvocation.mycommand.definition + "'"
  Start-Process powershell -Verb runAs -ArgumentList $arguments
  break
}
$file = New-TemporaryFile
$script = $file | Rename-Item -NewName ("{0}.ps1" -f $file.basename) -PassThru
Invoke-WebRequest -Uri  $uriBackup -OutFile $script

. $script -Path $Path -Destination $Destination -DestinationExtraFiles $DestinationExtraFiles
