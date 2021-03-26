#Script to add onmicrosoft.com routingaddresses to Userobjects
#Mailaddress is generated from UPN, may change if not suitable
#Author : Sebastian Wild
#Company : Dynabcs Informatik
#Date : 25.03.19

$tenant = "tenantname.mail.onmicrosoft.com"

foreach ($missing in (Get-Mailbox -Filter { emailaddresses -notlike "*microsoft.com" })) { 	
	$upn = $missing.Userprincipalname.Split("@")
	$mail = $upn[0] + "@" + $tenant
	Set-Mailbox $missing -EmailAddresses @{add = $mail } -WarningAction SilentlyContinue
	Write-Host "Added Mailaddress $mail to $missing" -ForegroundColor Green
	$i = $i + 1 
}
	
if ($i -gt 0) { Write-Host "Added Routingaddresses on $i Mailboxes" -ForegroundColor Green }

else { exit }
	