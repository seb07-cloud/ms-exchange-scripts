<#
.SYNOPSIS
    Generate new MoveRequests from AD Group Membership or OU
.DESCRIPTION
    Generate new MoveRequests from AD Group Membership or OU
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

function New-O365MoveRequest {
	[CmdletBinding()]
	param (
		[string]$Group,
		[string]$OU,
		[string]
	)
	
	begin {
		Import-Module ActiveDirectory
		Import-Module MSOnline
		try {

			$opensession = if (!(Get-PSSession -Name "ExchangeOnline*")) {

				$cred = Get-StoredCredential -target O365
				$opcred = Get-StoredCredential -target AD
	
			
				Connect-MsolService -Credential $cred
				$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
				$importresults = Import-PSSession $s -AllowClobber
			}
		}
		catch {
			Write-Host $_
		}
	}
	
	process {
		try {
			$endpoint = Get-MigrationEndPoint
			$mailboxes = Get-Mailbox | Select-Object UserPrincipalName

			if ($Group) {
				$users = Get-ADGroupMember -Identity $Group | Select-Object SamAccountName, objectClass | Where-Object { ( $_.SamAccountName -ne "ÖffentlicherOrdner" ) -and ( $_.SamAccountName -notlike "admin*" ) -and ( $_.SamAccountName -notlike "Mailbox*" ) -and ( $_.objectClass -eq "user" ) } | ForEach-Object { Get-ADUser $_.SamAccountName | Select-Object userPrincipalName }

				Foreach ($user in $users) {
					if ($user -notin $mailboxes.UserPrincipalName) {
						New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain 'hlbv365.mail.onmicrosoft.com' -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true | Out-Null
						Write-Host 'MoveRequest für ' $user ' erstellt' -ForeGroundColor Green
					}
					else { Write-Host "Mailbox for User $user already exists !" -ForeGroundColor Red }
				}
			}
			else {
				$users = Get-ADUser -SearchBase $ou -Properties mail -Filter { mail -like '*' } | `Select-Object Name, UserPrincipalName, Mail 

				Foreach ($user in $users) {
					if ($user -notin $mailboxes.UserPrincipalName) {
						New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain 'hlbv365.mail.onmicrosoft.com' -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true | Out-Null
						Write-Host 'MoveRequest für ' $user ' erstellt' -ForeGroundColor Green
					}
					else { Write-Host "Mailbox for User $user already exists !" -ForeGroundColor Red }
				}
			}
		}

		catch {
			Write-Host 'Fehler bei ' $user -ForeGroundColor Red
			Write-Host $_
		}
	}

	end {
		$confirm = Read-Host "Wollen sie die Exchange Online Session trennen? [y/n]"
		if ($confirm -eq 'y') {
			Get-PSSession $importresults | Remove-PSSession
		}
		else { exit }
	}
}

		
	
