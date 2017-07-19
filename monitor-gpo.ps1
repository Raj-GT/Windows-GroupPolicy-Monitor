<#
.SYNOPSIS
    Watch for Group Policy changes under monitored OU (and child OUs) and take automatic backups and optionally, alert via e-mail

.DESCRIPTION
    When run (ideally on a recurring schedule via Task Scheduler) the script will check Group Policies linked under $watchedOU for changes and perform an automatic backup of just the changed policies. It will also generate individual HTML/XML reports of the changed policies and save it with the backups. You can also have a summary of changes sent to you via e-mail.

    Each set of backup is created under it's own folder and kept indefinitely.
    
.INPUTS
    None

.OUTPUTS
    None

.LINK
    https://github.com/Raj-GT/Windows-GroupPolicy-Monitor

.NOTES    
    Version:    1.1
    Author:     Nimal Raj
    Revisions:  19/07/2017      Initial draft of v1.1

#>

#Requires -Version 3.0

#---------------------------------------------------------[Modules]---------------------------------------------------------
Import-Module ActiveDirectory,GroupPolicy

#--------------------------------------------------------[Variables]--------------------------------------------------------
$watchedOU = "DC=CORP,DC=CONTOSO,DC=COM"
$SMTP = "relay.contoso.com"
$alertRecipient = ""      # Leave empty to skip e-mail alerts
$scriptPath = $PSScriptRoot                         # Change the default backup path if required
$reportType = "HTML"                                # Valid options are HTML and XML
$backupFolder = "$scriptPath\Backups\" + (get-date -Format "yyyy-MM-ddThhmmss")

# E-mail template
$mailbody = @'
<style>
    body,p,h3 { font-family: calibri; }
    h3  { margin-bottom: 5px; }
    th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
    td  { padding: 5px 20px; }
    tr  { background: #E7FFF9; }
</style>

<p>GPO Monitor has detected the following changes under watched OU</p>

#mailcontent#

'@

# No user variables beyond this point
$ErrorActionPreference = "SilentlyContinue"
$GPCurrent = @()

#--------------------------------------------------------[Functions]--------------------------------------------------------
Function Backup ($GPO) 
    {
        If ($GPO) {
            New-Item -Path $backupFolder -ItemType Directory;
            $GPO | Backup-GPO -Path $backupFolder;
            $GPO | ForEach-Object {Get-GPOReport -Name $_.PolicyName -ReportType $reportType -Path "$backupFolder\"+$_.PolicyName+".$reportType"};
        }
    }

Function DN2Canon ($OUPath)
    {
        $Canon = $OUPath.Replace("DC=CORP,DC=CONTOSO,DC=COM","CORP.CONTOSO.COM").Replace("OU=","").Split(",")
        [Array]::Reverse($Canon)
        $Canon = $Canon -join "\"
        return($Canon)
    }

#--------------------------------------------------------[Execution]--------------------------------------------------------
# Generate list of GPOs linked under $watchedOU
$GPOLinks = (Get-ADOrganizationalUnit -SearchBase $watchedOU -Filter 'gpLink -gt "*"' | Get-GPInheritance).gpolinks

ForEach ($GPO in $GPOLinks) {
    $GPCurrent += New-Object -TypeName PSCustomObject -Property @{
    PolicyName  = (Get-GPO $GPO.GpoId).DisplayName;
    UpdateTime  = (Get-GPO $GPO.GpoId).ModificationTime;
    Enabled     = $GPO.Enabled;
    Guid        = $GPO.GpoId;
    OU          = DN2Canon($GPO.Target); 
    }
}

# Load the list of GPOs from last run for comparison
$GPLast = Import-Clixml -Path "$scriptPath\GPLast.xml"

# If no list is available then assume first run, create the list, backup all GPOs under $watchedOU and generate HTML/XML reports
If (!$GPLast -AND $GPCurrent) {
    $GPCurrent | Export-Clixml -Path "$scriptPath\GPLast.xml";
    Backup($GPCurrent)
}
Else {
# Let's compare the old list ($GPLast) to the current one ($GPCurrent)

    $GPList = $GPCurrent
    # Check for GPOs removed (guid missing from the current list)
    $RemovedGPO = Compare-Object $GPLast $GPList -Property Guid -PassThru | Where-Object {$_.SideIndicator -eq "<="}

    # Check for new GPOs (new guid in the current list)
    $NewGPO = Compare-Object $GPLast $GPList -Property Guid -PassThru | Where-Object {$_.SideIndicator -eq "=>"}
    # Remove the new GPO from the list before checking for changes (since new == change)
    $GPList = Compare-Object $GPList $NewGPO -Property Guid -PassThru

    # Check for changed GPOs
    $ChangedGPO = Compare-Object $GPLast $GPList -Property UpdateTime -PassThru | Where-Object {$_.SideIndicator -eq "=>"}

    # If anything has changed then create a backup (of new and changed GPOs) and update the GPLast.xml list
    If ($RemovedGPO -OR $NewGPO -OR $ChangedGPO) {
        $GPCurrent | Export-Clixml -Path "$scriptPath\GPLast.xml" -Force;
        Backup($NewGPO);
        Backup($ChangedGPO);
    }

    # If $alertRecipient is not empty, then generate and send a summary of changes via e-mail
    If ($alertRecipient) {
        # Generate HTML tables for the report
        If ($NewGPO) { $reportbody += $NewGPO | ConvertTo-Html -Fragment -Property PolicyName,OU,UpdateTime -PreContent "<h3>Policies Added</h3>" }
        If ($ChangedGPO) { $reportbody += $ChangedGPO | ConvertTo-Html -Fragment -Property PolicyName,OU,UpdateTime -PreContent "<h3>Policies Updated</h3>" }
        If ($RemovedGPO) { $reportbody += $RemovedGPO | ConvertTo-Html -Fragment -Property PolicyName,OU,UpdateTime -PreContent "<h3>Policies Removed</h3>" }

        $mailbody = $mailbody.Replace("#mailcontent#",$reportbody)
        $mailbody = $mailbody.Replace("PolicyName","Policy Name")
        $mailbody = $mailbody.Replace("OU","Organizational Unit")
        $mailbody = $mailbody.Replace("UpdateTime","Update Time")

        Send-MailMessage -SmtpServer $SMTP -To $alertRecipient -From "donotreply@contoso.com" -Subject "Group Policy Monitor" -Body $mailbody -BodyAsHtml
    }
    
}