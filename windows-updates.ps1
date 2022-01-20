## ------------------------------------------------------------------
## PowerShell Script To Automate Windows Update
## Script should be executed with "Administrator" Privilege
## 
## Based off the work of rfduarte Github: https://github.com/rfduarte
##
## Updated to check for needed reboots and perform as needed with the 
## ability to skip automatic reboots at the command line. Some other
## enhancements to make it useful on schedule.
## 
## Michael Moro https://github.com/mrmichaelmoro
## ------------------------------------------------------------------

Param(
    
[Parameter(Mandatory=$false)]

[Switch] $NoReboot

) #end param

If ($Error) {
	$Error.Clear()
}
$Today = Get-Date

$UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$Session = New-Object -ComObject Microsoft.Update.Session

## Default value settings
$MyPath = $MyInvocation.MyCommand.Path | Split-Path -parent
$ReportFile = $MyPath +"\" + $Env:ComputerName + "_Report.txt"
$ErrorActionPreference = "SilentlyContinue" 

if ($NoReboot) {
	$Rebootable = $false
} else {
	$Rebootable = $true
}

## Reboot pending tests
$pendingRebootTests = @(
    @{
        Name = 'RebootPending'
        Test = { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'  -Name 'RebootPending' -ErrorAction Ignore }
        TestType = 'ValueExists'
    }
    @{
        Name = 'RebootRequired'
        Test = { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update'  -Name 'RebootRequired' -ErrorAction Ignore }
        TestType = 'ValueExists'
    }
    @{
        Name = 'PendingFileRenameOperations'
        Test = { Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Ignore }
        TestType = 'NonNullValue'
    }
)

## RebootCheck function - will run through checks above and set switches as needed.
function RebootCheck {
    Add-Content $ReportFile "`r`n"
    Write-Host "`t Checking for Post-Install Required Reboots ..." -ForegroundColor "Yellow"
	Add-Content $ReportFile "Checking for Post-Install Required Reboots"
	Add-Content $ReportFile "------------------------------------------------`r"
    Add-Content $ReportFile "`r`n"

	## Initialize
	$RebootPending = $false 

    foreach ($test in $pendingRebootTests) {
        $test_result = Invoke-Command -ScriptBlock $test.Test
        
        if ($test.TestType -eq 'ValueExists' -and $test_result) {
            $RebootPending = $true
        } 

        if ($test.TestType -eq 'NonNullValue' -and $test_result -and $test_result.($test.Name)) {
            $RebootPending = $true
        }
    }

	return $RebootPending
} 

## Initialize reporting
If (Test-Path $ReportFile) { Remove-Item $ReportFile }
New-Item $ReportFile -Type File -Force -Value "Windows Update Report For Computer: $Env:ComputerName`r`n" | Out-Null
Add-Content $ReportFile "Report Created On: $Today`r"
Add-Content $ReportFile "==============================================================================`r`n"

## Main engine start
Write-Host
Write-Host "`t Initializing and Checking for Applicable Updates. Please wait ..." -ForeGroundColor "Yellow"
$Result = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

If ($Result.Updates.Count -EQ 0) {
	Write-Host "`t There are no applicable updates for this computer." -ForeGroundColor "Green"
	Add-Content $ReportFile "There are no applicable updates for this computer.`r"
}
Else {

    ## Prepare the list of applicable updates
	Write-Host "`t Preparing List of Applicable Updates For This Computer ..." -ForeGroundColor "Yellow"
	Add-Content $ReportFile "List of Applicable Updates For This Computer`r"
	Add-Content $ReportFile "------------------------------------------------`r"
	For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
		$DisplayCount = $Counter + 1
    		$Update = $Result.Updates.Item($Counter)
		$UpdateTitle = $Update.Title
		Add-Content $ReportFile "`t $DisplayCount -- $UpdateTitle"
	}

    ## Initialize downloads of applicable updates
	$Counter = 0
	$DisplayCount = 0
	Add-Content $ReportFile "`r`n"
	Write-Host "`t Initializing Download of Applicable Updates ..." -ForegroundColor "Yellow"
	Add-Content $ReportFile "Initializing Download of Applicable Updates"
	Add-Content $ReportFile "------------------------------------------------`r"
	$Downloader = $Session.CreateUpdateDownloader()
	$UpdatesList = $Result.Updates
	For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
		$UpdateCollection.Add($UpdatesList.Item($Counter)) | Out-Null
		$ShowThis = $UpdatesList.Item($Counter).Title
		$DisplayCount = $Counter + 1
		Add-Content $ReportFile "`t $DisplayCount -- Downloading Update $ShowThis `r"
		$Downloader.Updates = $UpdateCollection
		$Track = $Downloader.Download()
		If (($Track.HResult -EQ 0) -AND ($Track.ResultCode -EQ 2)) {
			Add-Content $ReportFile "`t Download Status: SUCCESS"
		}
		Else {
			Add-Content $ReportFile "`t Download Status: FAILED With Error -- $Error()"
			$Error.Clear()
			Add-content $ReportFile "`r"
		}	
	}

    ## Initiate installation of applicable updates
	$Counter = 0
	$DisplayCount = 0
	Write-Host "`t Starting Installation of Downloaded Updates ..." -ForegroundColor "Yellow"
	Add-Content $ReportFile "`r`n"
	Add-Content $ReportFile "Installation of Downloaded Updates"
	Add-Content $ReportFile "------------------------------------------------`r"
	$Installer = New-Object -ComObject Microsoft.Update.Installer
	For ($Counter = 0; $Counter -LT $UpdateCollection.Count; $Counter++) {
		$Track = $Null
		$DisplayCount = $Counter + 1
		$WriteThis = $UpdateCollection.Item($Counter).Title
		Add-Content $ReportFile "`t $DisplayCount -- Installing Update: $WriteThis"
		$Installer.Updates = $UpdateCollection
		Try {
			$Track = $Installer.Install()
			Add-Content $ReportFile "`t Update Installation Status: SUCCESS"
		}
		Catch {
			[System.Exception]
			Add-Content $ReportFile "`t Update Installation Status: FAILED With Error -- $Error()"
			$Error.Clear()
			Add-content $ReportFile "`r"
		}	
	}

    ## Initiate reboot check post-updates
    $RebootPending = RebootCheck

    ## If reboot needed, restart the instance.
    if ($RebootPending) {
        Add-Content $ReportFile "`t Reboot required to complete updates. Rebooting..."
        Write-Host "`t`t -- Reboot required to complete updates. Rebooting..." -ForegroundColor "Red"
		## Unless explicitly declined at command line, server will always reboot if changes are pending.
		if ($Rebootable) {
			Restart-Computer -Force
		} else {
			Add-Content $ReportFile "`t Reboot has been deferred per command args. Skipping..."
			Write-Host "`t`t -- Reboot has been deferred per command args. Skipping..." -ForegroundColor "Yellow"
		}
    } else {
        Add-Content $ReportFile "`t Reboot is not required to complete updates."
        Write-Host "`t`t -- Reboot is not required to complete updates." -ForegroundColor "Green"
    }  

} 

