<#PSScriptInfo

.VERSION 1.3

.GUID a9a9268e-acf3-4972-8c29-a7480f409e63

.AUTHOR Nimal Raj

.TAGS GroupPolicy,Automation

.PROJECTURI https://github.com/Raj-GT/Windows-GroupPolicy-Monitor

.EXTERNALMODULEDEPENDENCIES ActiveDirectory,GroupPolicy

#>

<#
.SYNOPSIS
    Watch for Group Policy changes under monitored OU (and child OUs) and take automatic backups and optionally, alert via e-mail

.DESCRIPTION
    When run (ideally on a recurring schedule via Task Scheduler) the script will check Group Policies linked under $watchedOU for changes and perform an automatic backup of new and changed policies. It will also generate individual HTML/XML reports of the policies and save it with the backups with an option to send a summary of changes via e-mail.

    Each set of backup is created under it's own folder and kept indefinitely.
    
.INPUTS
    None

.OUTPUTS
    None

.LINK
    https://github.com/Raj-GT/Windows-GroupPolicy-Monitor

.LINK
    https://www.experts-exchange.com/articles/30751/Automating-Group-Policy-Backups.html

.NOTES    
    Version:    1.3
    Author:     Nimal Raj
    Revisions:  19/07/2017      Initial draft of v1.1
                20/07/2017      Published in PowerShell Gallery
#>

#Requires -Version 3.0

#---------------------------------------------------------[Modules]---------------------------------------------------------
Import-Module ActiveDirectory,GroupPolicy

#--------------------------------------------------------[Variables]--------------------------------------------------------
$watchedOU = "DC=CORP,DC=CONTOSO,DC=COM"            # Root DN works as well
$rootDN = "DC=CORP,DC=CONTOSO,DC=COM"               # Required for our quick and dirty DN2Canonical function
$domainname = "CORP.CONTOSO.COM"                    # Required for our quick and dirty DN2Canonical function
$SMTP = "relay.contoso.com"                         # Assumes port TCP/25. Add -Port to Send-MailMessage if different
$mailFrom = "donotreply@contoso.com"                # From-address. For authenticated relays add -Credential to Send-MailMessage
$alertRecipient = "windows-admins@contoso.com"      # To-address. Leave empty to skip e-mail alerts
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

<p>GPO Monitor has detected the following changes...</p>

#mailcontent#

'@

# No user variables beyond this point
$ErrorActionPreference = "SilentlyContinue"
$GPCurrent = @()
$GPLast = $null
$reportbody = $null

#--------------------------------------------------------[Functions]--------------------------------------------------------
Function Backup ($GPO) 
    {
        If ($GPO) {
            New-Item -Path $backupFolder -ItemType Directory;
            $GPO | Backup-GPO -Path $backupFolder;
            $GPO | ForEach-Object { Get-GPOReport -Name $_.PolicyName -ReportType $reportType -Path ("$backupFolder\"+$_.PolicyName+".$reportType") };
        }
    }

Function DN2Canon ($OUPath)
    {
        $Canon = $OUPath -Replace($rootDN,$domainname) -Replace("OU=","") -Split(",")
        [Array]::Reverse($Canon)
        $Canon = $Canon -Join "\"
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

    # If anything has changed then create a backup (of new and changed GPOs), update GPLast.xml list and send -email
    If ($RemovedGPO -OR $NewGPO -OR $ChangedGPO) {
        $GPCurrent | Export-Clixml -Path "$scriptPath\GPLast.xml" -Force;
        Backup($NewGPO);
        Backup($ChangedGPO);
    
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

            Send-MailMessage -SmtpServer $SMTP -To $alertRecipient -From $mailFrom -Subject "Group Policy Monitor" -Body $mailbody -BodyAsHtml
        }
    }
}