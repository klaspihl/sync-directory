<#
.SYNOPSIS
    Sync directories and put extra/old files in other folder structure.
    Extra files older then x days is purged from the "extra" folder.

    Uses robocopy
.DESCRIPTION
    Steps;
        1. Removes empty directories and files older then defined days
        2. Moves files on destination not existing on source to another directory
        4. Move earlier versions of files with added date in filename
        3. Syncs files, do not purge extra files.

    Logging to files;
    Date_Old.log
    Date_Extra.log
    Date_Sync.log
    Date_Error.log

.EXAMPLE
    Sync documents on fileserver to External disk D:
    .\sync-directory.ps1 -Path \\pihl-fs\Pihl\ -Destination D:\Backup\Documents\ -DestinationExtraFiles D:\Backup\BackupPihl-ExtraFiles\ -Verbose

.PARAMETER Path
    Source directory

.PARAMETER Destination
    Target directory

.PARAMETER DestinationExtraFiles
    Target directory of file and folders removed from source since previous sync

.PARAMETER PurgeAge
    Age of files before removed from 'extra' folder

.PARAMETER LoggingPath
    Logging directory

.OUTPUTS
    Logs to Eventlog, files and console
.NOTES

    If source path is empty or wrong target all files in destiantion is moved to the 'extra files' directory.
    Requires to be run as administrator for robocopys 'audit' function to work on getting extra files.
    2021-02-08 Version 0.9 Proof of concept with major bug.
    2022-04-10 Version 1.0 Working backup of files.

.LINK
    https://github.com/KlasPihl
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    $Path,
    [Parameter(Mandatory=$true)]
    $Destination,
    $DestinationExtraFiles= (Join-Path (Get-Item $Destination | Select-Object -ExpandProperty Parent | Select-Object -ExpandProperty fullname) "BackupPihl-ExtraFiles" ),
    [int]$PurgeAge=180,
    $LoggingPath=(Join-Path $env:ALLUSERSPROFILE 'BackupPihl')
)
#Requires -RunAsAdministrator
#Requires -Modules Microsoft.PowerShell.Management
function test-Directory {
    <#
    .SYNOPSIS
        Validates that logging directory exist and that compression is as requested
    .PARAMETER Path
        Path of logging directory to validate

    .PARAMETER Compress
        If compression of folder should be enabled or disabled
    .NOTES

        2021-02-08 Version 1 Klas.Pihl@Gmail.com
        2022-04-10 Version 2 Use reference parameter to add '\' if not present to input variable.
    #>
    [CmdletBinding()]
    param (
        [ref]$Path,
        [switch]$Compress
    )
    try {
        if(-not (Test-Path $Path.Value)) {
            new-item -ItemType Directory $Path.Value
        }
        $Path.Value = Join-Path $Path.Value $null #add backspace to path
        Write-Verbose "Validate folder and compression status '$Compress' on '$($Path.Value)'"
        $FolderStatus = Get-CimInstance -Query "SELECT * FROM CIM_Directory WHERE Name = '$($Path.Value.trimend('\').Replace('\','\\'))'"
        if($FolderStatus.Compressed -ne $Compress) {
            Write-Verbose "Changing compress to $Compress on folder $($Path.Value)"
            $CompressionState  = switch ($Compress) {
                $true {'Compress'}
                $false {'Uncompress'}
            }
            $FolderStatus | Invoke-CimMethod -MethodName $CompressionState | Out-Null
        }
    }
    catch {
        throw "Can not validate that ($path.Value) exist and have valid compression status of $Compress"
    }

}
function Invoke-LogSource {
    $log = "Application"
    $source = "BackupPihl"
    try
    {
        if ([System.Diagnostics.EventLog]::SourceExists($source) -eq $false)
        {
            write-host "Creating event source $source on event log $log"
            Try
            {
                [System.Diagnostics.EventLog]::CreateEventSource($source, $log)

                #Verify that the Source was really created, if the Source still don't exist, exit script
                if ([System.Diagnostics.EventLog]::SourceExists($source) -eq $false)
                {
                    Write-Error "Failed to create a new eventlog Source, please run the script with a user that has local administrator rights."
                }
                else
                {
                    Write-Host -foregroundcolor Green "Event source $source created."
                    Write-Log -Text "Event source $source created." -Type Information
                    return $null
                }
            }
            catch
            {
                Write-Error "Failed to create a new eventlog Source, please run the script with a user that has local administrator rights."
            }
        }
        else
        {
            #write-host -foregroundcolor yellow "Warning: Event source $source already exists. Cannot create this source on Event log $log"
            return $null
        }
    }
    catch
    {
        Write-Error "Could not find the Source: $Source"
    }
    exit 1
}
function Write-Log {
    Param
    (
        [Parameter(Mandatory=$True)][string]$Text,
        [Parameter(Mandatory=$False)][System.Diagnostics.EventLogEntryType]$Type,
        [Parameter(Mandatory=$False)][int]$EventID=0
    )
    if($Text -ne "")
    {
        if ($null -eq $Type)
        {
            $Type = [System.Diagnostics.EventLogEntryType]::information
        }
        #write-host "$Text"
        $Text="`n$Text"
        Write-EventLog -LogName "Application" -Source "BackupPihl" -Message "$Text" -EventId $EventID -EntryType $Type
    }
}

#region Main body
$Error.Clear()
Write-Host -ForegroundColor White "Start backup of $path with destination $Destination"

Invoke-LogSource
if ($PurgeAge -lt 1) {
    Write-Warning "Age of file before permanently removed from archive $DestinationExtraFiles is set to $PurgeAge"
    Pause
}
if(-not (Test-Path $Path)) {
    Write-Log -Type Error -Text "Can not validate source path '$path'"
    Write-Error "Can not validate source path '$path', exit"
    exit 1
}
test-Directory -Path ([ref]$LoggingPath) -Compress
test-Directory -Path ([ref]$Destination)
test-Directory -Path ([ref]$DestinationExtraFiles)


#region clear old archive files

Write-host -ForegroundColor Gray "Remove file older the $PurgeAge days from $DestinationExtraFiles"
$LogFile = Join-Path $LoggingPath ("{0}_removed.log" -f (get-date -format "yyyy-MM-dd"))
$AllOldArchiveFiles = Get-ChildItem -Path $DestinationExtraFiles -File -Recurse |
    Where-Object {
        $PSitem.LastWriteTime -lt (get-date).AddDays(-$PurgeAge) -and
        $PSitem.CreationTime -lt (get-date).AddDays(-$PurgeAge)
    }

    foreach ($File in $AllOldArchiveFiles) {
            try {
            $File | remove-item -force -ErrorAction Stop -Confirm:0
        } catch {
            $ErrorRemoveArchive++
            #Write-Error $error[0].Exception
            $Error.remove($Error[0])
            Write-Error "Failed remove $($File.FullName)"
        }
    }


    Write-host -ForegroundColor Gray "Remove empty folders from $DestinationExtraFiles"
    $AllOldDirectorys = Get-ChildItem -Path $DestinationExtraFiles -Directory -Recurse |
        where-Object {
            -not $PSitem.GetFiles() -and
			-not $PSitem.GetDirectories()
        }

    foreach ($Directory in $AllOldDirectorys) {
        try {
            $Directory | remove-item -force -ErrorAction Stop
        } catch {
            $ErrorRemoveArchive++
            $Error.remove($Error[0])
            Write-Error "Failed remove $($Directory.FullName)"
        }
    }

if($AllOldArchiveFiles -or $AllOldDirectorys) {
    Out-File -FilePath $LogFile -Append -InputObject ($AllOldArchiveFiles | Select-Object FullName, LastWriteTime,CreationTime,Length)
    Out-File -FilePath $LogFile -Append -InputObject ($AllOldDirectorys | Select-Object FullName, LastWriteTime,CreationTime,Length)
}
#endregion clear old archive files

#region move extra files
    Write-host -ForegroundColor Gray "Move files that no longer exist at source from $Destination to $DestinationExtraFiles"
    $LogFile = Join-Path $LoggingPath ("{0}_extra.log" -f (get-date -format "yyyy-MM-dd"))
    $CommandExtra = "ROBOCOPY $path $Destination  /L /E /Copyall /XO /X /tee /njh /njs /np /ns /ndl /log+:$LogFile /XF thumbs.db"
    $ExtraFiles = Invoke-Expression $CommandExtra | Select-String -Pattern "EXTRA File" | ForEach-Object {
        $PSItem -split 'Extra file' | Select-Object -Last 1
    }
    if($LASTEXITCODE -eq 16) {
        Write-Log -Type Error -Text $Error
        exit 1
    }

    if($ExtraFiles) {
        $ExtraFiles = $ExtraFiles.trim()
        Write-Verbose "Found old files at $Destination, moving to $DestinationExtraFiles"
        $ResultMove =
            ForEach ($File in $ExtraFiles) {
                try {
                    $File | Move-Item -Destination $DestinationExtraFiles -Confirm:$false -PassThru -ErrorAction Stop
                } catch {
                    $ErrorMoveArchive++
                    $Error.remove($Error[0])
                    Write-Error "$File already exist on destination $DestinationExtraFiles"
                    #Write-Error $error[0].Exception

                }
            }
        $LogTextExtraFiles = "$($ResultMove.count) files moved from $Destination to $DestinationExtraFiles`nFiles no longer exist on source"
        Write-Log -Type Information -Text $LogTextExtraFiles
        Write-Verbose -Message $LogTextExtraFiles
    }



#endregion move extra files

#region move older versions of file with added timestamp #cryptovirus?
Write-host -ForegroundColor Gray "Move earlier versions of files at $Destination to $DestinationExtraFiles"

    $LogFile = Join-Path $LoggingPath ("{0}_changed.log" -f (get-date -format "yyyy-MM-dd"))
    $CommandExtra = "ROBOCOPY $path $Destination  /L /E /copy:DAT /XO /X /tee /njh /njs /np /ns /ndl /log+:$LogFile /XF thumbs.db"
    $NewerFiles = Invoke-Expression $CommandExtra | Select-String -Pattern "Newer" | ForEach-Object {
        $PSItem -split 'Newer' | Select-Object -Last 1
    }

    if($NewerFiles) {
        #$NewerFiles = $NewerFiles.trim()
        Write-Verbose "Move files that exist with changed data on destination to $DestinationExtraFiles"
        $ResultMove =
            ForEach ($File in $NewerFiles.trim()) {
                try {
                    $OlderVersionFile = get-item $File.Replace($Path,$Destination)
                    $DestinationFileName = Join-Path $DestinationExtraFiles ("{0}_{1}{2}" -f $OlderVersionFile.BaseName,(get-date -format FileDateTime),$OlderVersionFile.Extension)

                    #$FileObject = Get-Item $File -ErrorAction Stop
                    #$DestinationFileName = Join-Path $DestinationExtraFiles ("{0}_{1}{2}" -f $FileObject.BaseName,(get-date -format "yyyy-MM-dd"),$FileObject.Extension)
                    $OlderVersionFile | Move-Item -Destination $DestinationFileName -Confirm:$false -PassThru -ErrorAction Stop -Force
                    Write-Log -Type Information -Text "Earlier version of file '$File' found at $Destination, move old version.`nNew name: $DestinationFileName"
                } catch {
                    $ErrorMoveArchive++
                    $Error.remove($Error[0])
                    #Write-Error "$DestinationFileName already exist"
                    Write-Error $error[0].Exception

                }
            }
        $LogTextExtraFiles = "$($ResultMove.count) files on $path existed on $Destination, moved earlier versions $DestinationExtraFiles"
        Write-Log -Type Warning -Text $LogTextExtraFiles
        Write-Verbose -Message $LogTextExtraFiles
    }
#endregion move older versions of file with added timestamp

#region sync files
Write-host -ForegroundColor Gray "Mirror files at $path to $Destination"

$LogFile = Join-Path $LoggingPath ("{0}_sync.log" -f (get-date -format "yyyy-MM-dd"))
$CommandSync = "robocopy $Path $Destination /e /sec /r:2 /w:1 /log+:$LogFile /V /NP /DCOPY:T /COPY:DAT /mt:64 /tee /XF thumbs.db"
$ResultSync = Invoke-Expression $CommandSync
if($LASTEXITCODE -gt 3)
    {

        $ErrorSync = $ResultSync| Select-Object -Last 11 | Select-Object -first 9 | Out-String
        Write-Error "`n$ErrorSync)"
        Write-Error "Lastexitcode: $LASTEXITCODE"
    }

#endregion sync files

if($ErrorRemoveArchive -or $ErrorMoveArchive -or $ErrorSync) {
    $LogFile = Join-Path $LoggingPath ("{0}_error.log" -f (get-date -format "yyyy-MM-dd"))
    Out-File -FilePath $LogFile -Append -InputObject ($Error | Out-String) -Encoding utf8
    Write-Log -Type Error -Text $Error
} else {
    Write-Log -Type Information -Text "Successful sync on $path to $Destination`n$($ResultSync| Select-Object -Last 11 | Select-Object -first 9 | Out-String)"
    Write-Host -ForegroundColor White "Successful sync on $path to $Destination`n$($ResultSync| Select-Object -Last 11 | Select-Object -first 9 | Out-String)"
}

#endregion Main body

