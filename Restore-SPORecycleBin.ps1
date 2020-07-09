param (
  [Parameter(Mandatory=$true)][String]$Organisation,
  [Parameter(Mandatory=$true)][String]$MSTeamsSitesCSV,
  [Parameter(Mandatory=$true)][String]$CSVReport,
  [Parameter(Mandatory=$false)][bool]$RestoreFiles = $false
)

<#
.SYNOPSIS
  Script will restore files from recycle bin from SharePoint Online sites.

.DESCRIPTION
  This script will restore all files which are deleted by the System Account. These files
  was deleted because of a wrong configured deletion policy.

.COMPONENT
  Requires PowerShell module:
  - Microsoft.Online.SharePoint.Powershell
  - SharePointPnPPowerShellOnline

.PARAMETER Organisation
  This parameter is mandatory and defined the Microsoft 365 tenant name.

.PARAMETER MSTeamsSitesCSV
  This parameter is mandatory and provide the full path to the CSV file in which all SharePoint
  Online sites are listed to check the recycle bin for restoring.

.PARAMETER CSVReport
  This parameter is mandatory and provide the full path to store the report of all items which
  are identified in the recycle bin for restoring.

.PARAMETER RestoreFiles
  This parameter is optional and the default value is $false. Only by $true the restore will be
  really done.

.NOTES
  Version:          1.0
  Author:           Stefan Gericke
  Creation Date:    2020/07/09
  Description:      Script created

.EXAMPLE
  Restore-SPORecycleBin -Organisation <Org Name> -MSTeamsSitesCSV <Path of CSV file with all SPO sites to check> -CSVReport <Path of the report> -RestoreFiles $true
#>

#----------------------------------------------[Functions]-----------------------------------------------------

function New-LogFile {
    param (
      [Parameter(Mandatory=$true)][string]$Path,
      [Parameter(Mandatory=$true)][string]$FileName,
      [Parameter(Mandatory=$false)][string]$TimeStamp = "",
      [Parameter(Mandatory=$false)][string]$Step = ""
    )
  <#
  .SYNOPSIS
  This function will return the path.
  
  .DESCRIPTION
  This function will return the path for a logfile which can be used for log files. If you don't get a timestamp
  in the parameter this function will generate a timestamp for the filename.
  
  .OUTPUTS
  The complete path of the log file.
  
  .PARAMETER Path
  The path of the folder where you want to store the log file.
  
  .PARAMETER FileName
  The file name you want to have included after the timestamp.
  
  .PARAMETER TimeStamp
  If you already have a timestamp you want to use you can give this optional.
  
  .PARAMETER Step
  This is the step of the migration process to include it in the file name.
  #>
  
    if (!(Test-Path -Path $Path)) { New-Item -Path $Path -ItemType Directory -Force }
    if ($TimeStamp -eq "") {
        $TimeStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        if ($Step -ne "") {
            $PathLogFile = $Path + $TimeStamp + "_" + $FileName + "_" + $Step + ".log"
        } else {
            $PathLogFile = $Path + $TimeStamp + "_$FileName.log"
        }
    } else {
        if ($Step -ne "") {
            $PathLogFile = $Path + $TimeStamp + "_" + $FileName + "_" + $Step + ".log"
        } else {
            $PathLogFile = $Path + $TimeStamp + "_$FileName.log"
        }
    }
    return $PathLogFile
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$LogMessage,
        [string]$LogFile,
        [ValidateSet("Error","Warn","Info","Ok","Failed","Success")][string]$LogLevel
    )
    $arr = @()
    
    # Set Date/Time
    $DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $arr += $DateTime
    # Set log level
    switch ($LogLevel) {
        'Error' {
            $arr += 'ERROR    '
        }
        'Warn' {
            $arr += 'WARN     '
        }
        'Info' {
            $arr += 'INFO     '
        }
        'Ok' {
            $arr += 'OK       '
        }
        'Failed' {
            $arr += 'FAILED   '
        }
        'Success' {
            $arr += 'SUCCESS  '
        }
        Default {
            $arr += '         '
        }
    } # END SWITCH
        
    # Set message
    $arr += $LogMessage

    # Build line from array
    $line = [System.String]::Join(" ",$arr)

    # Write to log
    If ($LogFile -ne "") { $line | Out-File -FilePath $LogFile -Append }
                
    # Write to host
    Write-Host $line
}

#----------------------------------------------[Declarations]-----------------------------------------------------

# Variables
$SPOAdminURL = "https://$Organisation-admin.sharepoint.com"
$logFileJob = New-LogFile -Path "./" -FileName "RestoreRecycleBin"
$cred = Get-Credential -Message "SharePoint Online Administrator"

# IMPORTANT: Needed for the Federation Server which speaks only TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#-----------------------------------------------[Execution]-------------------------------------------------------

# Import list of MS Teams site from CSV file
try {
    $list = Import-Csv -Path $MSTeamsSitesCSV -Delimiter ";" -Header SharePointUrl
    Write-Log -LogMessage "CSV file is loaded ..." -LogFile $logFileJob
    $MSTEamsSites = @()
    foreach ($item in $list) {
        $MSTEamsSites += $item.SharePointUrl
    }
}
catch {
    Write-Log -LogMessage "Failure by importing the CSV file $CsvFile" -LogLevel Error -LogFile $logFileJob
    Write-Log -LogMessage $_.Exception.Message -LogLevel Error -LogFile $logFileJob
    Write-Log -LogMessage "*** Stop script Restore-SPORecycleBin ***" -LogLevel Error -LogFile $logFileJob
    exit
}

try {
    # Connect to the SPO Services to add and remove Service Account to Site Collection Administrator
    Connect-SPOService -Url $SPOAdminURL -Credential $cred
    Write-Log -LogMessage "Connected to SPO Services ..." -LogFile $logFileJob

    # Remove report file if existing
    if (Test-Path -Path $CSVReport) { Remove-Item -Path $CSVReport -Force }

    foreach ($item in $MSTEamsSites) {
        Write-Log "Starting with MS Teams site $item ..." -LogLevel Info -LogFile $logFileJob
        try {
            # Add Service Account to Site Collection Administrator
            Set-SPOUser -site $item -LoginName $ServiceAccount -IsSiteCollectionAdmin $true
            Write-Log -LogMessage "Service account is added to MS Teams site ..." -LogFile $logFileJob

            # Parameters
            $Account = "System Account"
            $DeletionDate = "07/02/2020"

            # Connect to the MS Teams site with PnP
            Connect-PnPOnline -Url $item -Credentials $cred
            Write-Log -LogMessage "Connected to Team Site $item with PnP ..." -LogFile $logFileJob
            
            # Identify the files to recycle
            $recycleBinItems = Get-PnPRecycleBinItem -FirstStage | Where-Object {($_.DeletedByName -eq $Account) -and ($_.DeletedDate -ge $DeletionDate) -and ($_.Title -notlike "*siteicon*")}
            #$recycleBinItems = Get-PnPRecycleBinItem | Where-Object {($_.DeletedDate -ge $DeletionDate)}
            Write-Log -LogMessage "$($recycleBinItems.Count) files are identified in the recycle bin!" -LogLevel Info -LogFile $logFileJob

            if ($RestoreFiles) { # Restore only when parameter is true
                # Foreach of each file under the Recycle bin
                foreach ($recycleBinItem in $recycleBinItems) {
                    try {
                        Write-Log -LogMessage "Try to restore $($recycleBinItem.DirName)/$($recycleBinItem.Title) ..." -LogFile $logFileJob
                        Restore-PnpRecycleBinItem -Identity $recycleBinItem -Force -ErrorVariable ErrVar

                        Write-Log -LogMessage "Try to restore $($recycleBinItem.DirName)/$($recycleBinItem.Title) ... Done!" -LogLevel Info -LogFile $logFileJob
                        Write-Log -LogMessage "Add $($recycleBinItem.DirName) $($recycleBinItem.Title) report file ..." -LogFile $logFileJob

                        if ($ErrVar) {
                            Write-Log -LogMessage "Try to restore $($recycleBinItem.DirName)/$($recycleBinItem.Title) ... Failed!" -LogLevel Error -LogFile $logFileJob
                            # Export information to CSV file
                            $recycleBinItem | Select-Object @{Name="SPOUrl";Expression={$item}},DirName,Title,AuthorName,DeletedDate,@{Name="Restored";Expression={"Failed"}} `
                            | Export-Csv $CSVReport -NoTypeInformation -Encoding UTF8 -Append
                        } else {
                            # Export information to CSV file
                            $recycleBinItem | Select-Object @{Name="SPOUrl";Expression={$item}},DirName,Title,AuthorName,DeletedDate,@{Name="Restored";Expression={"Successful"}} `
                            | Export-Csv $CSVReport -NoTypeInformation -Encoding UTF8 -Append
                        }
                    } catch {
                        Write-Log -LogMessage "Try to restore $($recycleBinItem.DirName)/$($recycleBinItem.Title) ... Failed!" -LogLevel Error -LogFile $logFileJob
                        # Export information to CSV file
                        $recycleBinItem | Select-Object @{Name="SPOUrl";Expression={$item}},DirName,Title,AuthorName,DeletedDate,@{Name="Restored";Expression={"Failed"}} `
                        | Export-Csv $CSVReport -NoTypeInformation -Encoding UTF8 -Append
                    }
                }  # foreach
            } else { # Restore recycle bin is false
                # Export identified items in recycle bin to report file
                $recycleBinItems | Select-Object @{Name="SPOUrl";Expression={$item}},DirName,Title,AuthorName,DeletedDate `
                  Export-Csv $CSVReport -NoTypeInformation -Encoding UTF8 -Append
            }
            Disconnect-PnPOnline
    
            # Remove Service Account to Site Collection Administrator
            Set-SPOUser -site $item -LoginName $ServiceAccount -IsSiteCollectionAdmin $false
            Write-Log -LogMessage "Service account is removed from the MS Teams site ..." -LogFile $logFileJob

            Write-Log "Finished with MS Teams site $item ..." -LogLevel Info -LogFile $logFileJob
        } catch {
            Write-Log -LogMessage $_.Exception.Message -LogLevel Error -LogFile $logFileJob
            Write-Log -LogMessage "*** MS Teams site $item failed! ***" -LogLevel Error -LogFile $logFileJob
        }
    }

    # Disconnect to services
    Disconnect-SPOService
    Write-Log -LogMessage "*** Stop script Restore-SPORecycleBin ***" -LogLevel Ok -LogFile $logFileJob
} catch {
    Write-Log -LogMessage $_.Exception.Message -LogLevel Error -LogFile $logFileJob
}
    Write-Log -LogMessage "*** Stop script Restore-SPORecycleBin ***" -LogLevel Error -LogFile $logFileJob