
<#
.SYNOPSIS
  An SCCM Maintenance Window is needing to be run shortly after every Patch Tuesday; for example, the Wednesday following.

  Patch Tuesday is every second Tuesday of the month, but that does NOT mean the next Wednesday will be the second Wednesday of the month.

  An Offset needs to be calculated, but this cannot yet be natively done in SCCM.

  This script provides an automated solution to the problem.

.DESCRIPTION
This script is intended to run the first of the month from SCCM server itself (scheduled task)
It assumes that the Collections to adjust maintenance windows for have a consistent naming pattern which can be used.
It assumes a manual,initial maintenance window has been created on the collection, with the appropriate Date or Time.

Each relevant collection will be discovered, have it's maintenance window assessed, and then updated to be the same day of the week after the next Patch Tuesday
    - For example:
        - if the maintenance window is on Wednesday @ 7:00PM, the updated value will be Patch Tuesday + 1 day @ 7:00PM
        - if the maintenance window is on Thursday @ 2:00AM, the updated value will be Patch Tuesday + 2 day @ 2:00AM

.INPUTS
    None
.OUTPUTS
  Log file stored in C:\Scripts\Logs\SCCM_UpdateMaintenanceWindowTime.txt
  Email delivered through SMTP
.NOTES
  Version:        1.0
  Author:         Jeff Miles
  Creation Date:  January 16, 2020

#>
Param (
    [string]$sitecodeid = "SDE", # Provide the SiteCode ID you want to analyze Collections for
    [string]$CollectionFilter = "MW*", # Provide the string that you filter collections on to find the ones for this scripts purpose
        # Wildcard is valid here, because the usage of this parameter allows for it.
        # For example, find all collections that begin with "MW"
    [string]$EmailDestination = "destination@domain.com"
)

function Send-InternalEmail {
    Param (
        [Parameter(Mandatory = $true)]
        [String]$EmailTo,
        [Parameter(Mandatory = $true)]
        [String]$Subject,
        [Parameter(Mandatory = $true)]
        [String]$Body,
        [Parameter(Mandatory = $false)]
        [String]$EmailFrom = 'email@domain.com',
        [parameter(Mandatory = $false)]
        [String] $SmtpServer = "relay.domain.com",
        [parameter(Mandatory = $false)]
        [String] $SmtpUsername,
        [parameter(Mandatory = $false)]
        [SecureString] $SmtpPassword
    )
    $SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom, $EmailTo, $Subject, $Body)
    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)
    $SMTPMessage.IsBodyHtml = $true
    $SMTPClient.Send($SMTPMessage)
} #End Function Send-EMail

# Logic of date calculation from here: https://www.madwithpowershell.com/2014/10/calculating-patch-tuesday-with.html
# First, find the 12th day of the current month
$BaseDate = ( Get-Date -Day 12 ).Date
# For example: January 12 2020 00:00:00

# 2 is the Tuesday date of week integer. Subtract the day of week of the 12th from 2 (Tuesday) to find the difference for AddDays.
# Since we run on the first of the month, assume that this will be the date we use
$PatchTuesday = $BaseDate.AddDays( 2 - [int]$BaseDate.DayOfWeek )
# For Example: [int]$BaseDate.DayOfWeek For January = 0 (Sunday) because January 12 2020 is Sunday
#  2 - 0 = 2, so if $BaseDate = Sunday Jan 12 and we add 2 days, we get Tuesday Jan 14th. Patch Tuesday!

# Just in case running out of normal schedule (not on the first of the month, perhaps AFTER patch tuesday of the current month)
if ( (Get-Date).Date -gt $PatchTuesday ) {
    # if today is greater than patch tuesday for the month
    # get next months' date
    $BaseDate = $BaseDate.AddMonths( 1 )
    $PatchTuesday = $BaseDate.AddDays( 2 - [int]$BaseDate.DayOfWeek )
}

$EmailCollection = @() # used to store strings for email body of what happened in the script.

#Import SCCM Module from what is expected to be a default path, using Environment Variable
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("bin\i386", "bin\configurationmanager\configurationmanager.psd1")

# Need this command in case script is running as system which doesn't load the PSDrive automatically
# https://stackoverflow.com/questions/42155927/powershell-cd-sitecode-not-working
New-PSDrive -Name $sitecodeid -PSProvider "AdminUI.PS.Provider\CMSite" -Root "$ENV:ComputerName" -Description "SCCM Site" -ErrorAction ignore

#Get SiteCode object
$SiteCode = Get-PSDrive -PSProvider CMSITE

# Set shell Location to SCCM PSDrive
Set-Location "$($SiteCode.Name):"

#Find all collections with MW in the name, but exclude the fake one
$Collections = Get-CMCollection -Name $CollectionFilter

# Exclude any Collections that you wish (not currently parameterized in this script)
$Collections = $Collections | Where-Object { $_.Name -notlike "MW - Fake*" -and $_.Name -notlike "MW*reoccurring" }


foreach ($col in $Collections) {
    # Need to loop through so we have a unique collection ID
    # Get the maintenance window on the collection
    $maintenancewindow = Get-CMMaintenanceWindow -CollectionID $col.CollectionID

    # Find existing day of week on that maintenance window, in integer form
    $existingday = [int]$maintenancewindow.StartTime.DayOfWeek

    # Assume existing day is wednesday (3) or thursday (4). This will add either 1 or 2 days to patch tuesday (future day of week minus tuesday day of week)
    $PatchDate = $PatchTuesday.AddDays($existingday - 2)
        # This results in the actual date we want to do the patching on

    # Ensure we set the PatchDate time as the same as existing time on the maintenance window (instead of 00:00:00 by default)
    $time = [System.Timespan]::Parse("$($maintenancewindow.StartTime.Hour):$($maintenancewindow.StartTime.Minute)")
    $PatchDate = $PatchDate.Add($time)

    try {
        # Define the new schedule, and make sure end time duration matches original
        $sched = New-CMSchedule -Start $PatchDate -End $PatchDate.AddMinutes($maintenancewindow.Duration) -Nonrecurring
        # Set the maintenance window. If this fails, then the Catch will trigger
        Set-CMMaintenanceWindow -Name $maintenancewindow.Name -CollectionID $col.CollectionID -Schedule $sched
        # Lets re-get the maintenance window as a confirmation in the log and email, instead of just a calculated value.
        $Updatedmaintenancewindow = Get-CMMaintenanceWindow -CollectionID $col.CollectionID

        $EmailCollection += (Get-Date -format "yyyy-MM-dd hh:mm:s UTC zz")
        $EmailCollection += "Modifying Maintenance Window: $($maintenancewindow.Name) <br />"
        $EmailCollection += "  - original date: $($maintenancewindow.StartTime) <br />"
        $EmailCollection += "  - new date: $($Updatedmaintenancewindow.StartTime) <br />"
        $EmailCollection += " -------------- <br /><br />"
    }
    catch {
        $EmailCollection += (Get-Date -format "yyyy-MM-dd hh:mm:s UTC zz")
        $EmailCollection += "An error occurred with Maintenance Window: $($maintenancewindow.Name) <br />"
        $EmailCollection += $_
        $EmailCollection += " -------------- <br /><br />"
    }
}
# Log output of the file
$fileresult = $EmailCollection -replace '<br />', "`r`n"
Add-Content C:\Scripts\Logs\SCCM_UpdateMaintenanceWindowTime.txt -Value $fileresult

# Send email with output of what happened.
$EmailCollection = $EmailCollection | out-string
if (!$EmailCollection) { $EmailCollection = "No SCCM Maintenance Window modifications occurred." }
Send-InternalEmail -EmailTo $EmailDestination -Body $EmailCollection -Subject "SCCM Maintenance Window modifications"