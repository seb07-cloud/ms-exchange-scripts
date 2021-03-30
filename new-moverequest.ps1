Import-Module ActiveDirectory
Import-Module MSOnline

# Get StoredCredential
$cred = Get-StoredCredential -target O365
$opcred = Get-StoredCredential -target AD

# Connect Msol

Connect-MsolService -Credential $cred
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
$importresults = Import-PSSession $s -AllowClobber


$endpoint = Get-MigrationEndPoint

# User aus Gruppe oder OU

#$users = Get-ADUser -SearchBase 'OU=Users,OU=Domain_Users,DC=Domain,DC=local' -Properties mail -Filter {mail -like '*'} |`Select-Object Name, UserPrincipalName, Mail 
$users = Get-ADGroupMember -Identity gr_O365-Sync | Where-Object { $_.SamAccountName -ne "ÖffentlicherOrdner" } | ForEach-Object {Get-ADUser $_.SamAccountName | Select-Object userPrincipalName }

$users | Export-CSV .\MigrationBatch.csv -NoTypeInformation

Foreach ($user in $users)
	{
	Try {
		New-MoveRequest -Erroraction Stop -Identity $user.userPrincipalName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain 'hlbv365.mail.onmicrosoft.com' -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true | Out-Null
		Write-Host 'MoveRequest für ' $user ' erstellt' -ForeGroundColor Green
		}
	Catch 
		{
		Write-Host 'Fehler bei ' $user -ForeGroundColor Red
		}
	}	
		
Get-PSSession | Remove-PSSession
