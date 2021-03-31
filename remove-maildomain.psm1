<#
.SYNOPSIS
    Removes specific Maildomains from AD User Objects
.DESCRIPTION
    Removes specific Maildomains from AD User Objects
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\New-O365MoveRequest -Group "gr_O365-Sync"
	.\New-O365MoveRequest -OU "OU=Contoso-Groups,DC=Contoso,DC=local"

.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   30.03.2021

    Changelog:
		1.0             Initial Release
#>

function Remove-MailDomain {
    param (
        [Parameter(Mandatory)]
        [string]$Domain
    )
    begin {}
    
    process {
        try {
            $users = get-mailbox | Where-Object { $_.emailaddresses -like $Domain }
            foreach ($user in $users) {
                $addresses = (get-mailbox $user.alias).emailaddresses
                $fixedaddresses = $addresses | Where-Object { $_.proxyaddressstring -notlike $Domain }
                set-mailbox $user.alias -emailaddresses $fixedaddresses
                Write-Host "Removed Maildomain from" $user.Name -ForeGroundColor Green
            }

        }
        catch {
            Write-Host $_
        }
    }
    end {}
}