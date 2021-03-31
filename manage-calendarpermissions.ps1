<#
.SYNOPSIS
    Script to add or remove calendar Permissions
.DESCRIPTION
    Script to add or remove calendar Permissions
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\manage-calendarpermissions.ps1

.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   22.03.2021
       
    Changelog:
		1.0             Initial Release
#>

function Edit-MailboxPermission {
    [CmdletBinding()]
    param (
        [ValidateSet("ExchangeOnline", "Exchange")]
        [Parameter(Mandatory)]
        [string]$ExchangeType
    )
    
    begin {

        $permissions = New-Object System.Collections.ArrayList 
        $permissions.Add("Owner") | Out-Null
        $permissions.Add("Editor") | Out-Null
        $permissions.Add("Reviewer") | Out-Null
        $permissions.Add("Author") | Out-Null

        if ($ExchangeType -eq "Exchange") {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 
            Start-Sleep 5
            $mailboxes = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name
        }

        elseif ($ExchangeType = "ExchangeOnline") {
            #Requires -Module MSOnline
            $cred = Get-Credential
            Import-Module MSOnline
            Connect-MsolService -Credential $cred
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
            $importresults = Import-PSSession $s -AllowClobber
            Start-Sleep 5
            $mailboxes = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name
        }

    }
    
    process {

        $error.Clear()    

        $gridmailboxtomanage = $mailboxes | Out-GridView -PassThru -Title "Wählen sie die Mailbox aus, auf die sie Berechtigungen vergeben möchten."
        $gridusertoadd = $mailboxes | Out-GridView -PassThru -Title "Wählen sie die Mailbox aus, welche Berechtigungen bekommen soll."
        $gridchosen = $permissions | Out-GridView -PassThru -Title "Wählen sie die Berechtigungsstufe aus."

        try {
            if (!(Get-MailboxfolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias)) {
                Add-MailboxFolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias -AccessRights $gridchosen 
            }
            else {
                Set-MailboxFolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias -AccessRights $gridchosen 
            }
        }
        catch {
            Write-Host "Trying to Pull Calendar in English instead of German ..." -ForegroundColor Yellow
        }

        if (!$error) {
            Write-Host "Mailboxberechtigung" $chosen "für" $gridusertoadd "wurde auf" $gridmailboxtomanage "korrekt gesetzt!" -ForegroundColor Green 
        }

        else {

            try {
                if (!(Get-MailboxfolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias)) {
                    Add-MailboxFolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias -AccessRights $gridchosen 
                }
                else {
                    Set-MailboxFolderPermission -Identity "$($gridmailboxtomanage.Alias):\Kalender" -User $gridusertoadd.Alias -AccessRights $gridchosen 
                }
            }
            catch {
                Write-Host "Couldnt set up Calendar Permission | $error[0] " -ForegroundColor Red
            }

            if (!$error) {
                Write-Host "Mailboxberechtigung" $gridchosen "für" $gridusertoadd "wurde auf" $gridmailboxtomanage "korrekt gesetzt!" -ForegroundColor Green 

            }
        }
    }
    end {}
}

