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
		[string]$OU
	)
	
	begin {
		Import-Module ActiveDirectory
		Import-Module MSOnline
		try {
			$cred = Get-StoredCredential -target O365
			$opcred = Get-StoredCredential -target AD
	
			
			Connect-MsolService -Credential $cred
			$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
			$importresults = Import-PSSession $s -AllowClobber
		}
		catch {
			Write-Host $_
		}
	}
	
	process {
		try {
			$endpoint = Get-MigrationEndPoint

			if ($Group) {
				$users = Get-ADGroupMember -Identity gr_O365-Sync | Where-Object { $_.SamAccountName -ne "ÖffentlicherOrdner" } | ForEach-Object { Get-ADUser $_.SamAccountName | Select-Object userPrincipalName }

				Foreach ($user in $users) {
					New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain 'hlbv365.mail.onmicrosoft.com' -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true | Out-Null
					Write-Host 'MoveRequest für ' $user ' erstellt' -ForeGroundColor Green
				}
			}
			else {
				$users = Get-ADUser -SearchBase $ou -Properties mail -Filter { mail -like '*' } | `Select-Object Name, UserPrincipalName, Mail 

				Foreach ($user in $users) {
					New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain 'hlbv365.mail.onmicrosoft.com' -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true | Out-Null
					Write-Host 'MoveRequest für ' $user ' erstellt' -ForeGroundColor Green
				}
			}
		}
		catch {
			Write-Host 'Fehler bei ' $user -ForeGroundColor Red
			Write-Host $_
		}

		end {
			Get-PSSession $importresults | Remove-PSSession
		}
	}
}

		
	
