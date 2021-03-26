Import-Module ActiveDirectory
Import-Module MSOnline

# Get On Prem Credentials

$opcred = Get-Credential domain\administrator

# Connect MSOnline

$cred = Get-Credential administrator@tenant.onmicrosoft.com
Connect-MsolService -Credential $cred
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
$importresults = Import-PSSession $s -AllowClobber


# Setzen Variablen

$grgerman = "SGUSR-Azure-Language-de"
$grenglish = "SGUSR-Azure-Language-en"
$grlizenzstandard = "gr_Standard"
$grlizenzenterprise = "gr_Enterprise"
$lizenzstandard = "Tenant:STANDARDPACK"
$lizenzenterprise = "Tenant:ENTERPRISEPACK"

# Importieren CSV 

$users = Import-Csv .\MigrationBatch.csv
$forwardimport = Import-Csv .\Mailforwarding.csv

# Setzen Mailbox Einstellungen

	Foreach ($user in $users)
		{
		
			# Gruppenzugeh�rigkeit Sprache
			
			$membersde = Get-ADGroupMember -Identity $grgerman -Recursive
			$membersen = Get-ADGroupMember -Identity $grenglish -Recursive
			
			If ($membersde.Name -contains $user.Name) 
				{
				$language = "DE-de"
				Write-Host "$user.UserPrincipalName in Gruppe $grgerman gefunden" -ForeGroundColor Green 
				} 
			ElseIf ($membersen.Name -contains $user.Name) 
				{
				$language = "EN-en"
				Write-Host "$user.UserPrincipalName in Gruppe $grenglish gefunden" -ForeGroundColor Green 
				}
			Else 
				{
				$language = $null
				Write-Host "$user.UserPrincipalName ist in keiner Sprach Gruppe" -ForeGroundColor Red
				}
	
			# Gruppenzugeh�rigkeit Lizenz
			
			$membersstd = Get-ADGroupMember -Identity $grlizenzstandard -Recursive
			$membersepr = Get-ADGroupMember -Identity $grlizenzenterprise -Recursive
			
			If ($membersstd.Name -contains $user.Name) 
				{
				$lizenz = $lizenzstandard
				Write-Host "$user.UserPrincipalName in Gruppe $grlizenzstandard gefunden" -ForeGroundColor Green 
				} 
			ElseIf ($membersepr.Name -contains $user.Name) 
				{
				$lizenz = $lizenzenterprise
				Write-Host "$user.UserPrincipalName in Gruppe $grlizenzenterprise gefunden" -ForeGroundColor Green 
				}
			Else 
				{
				Write-Host "$user.UserPrincipalName ist in keiner Lizenz Gruppe" -ForeGroundColor Red
				$lizenz = $null
				}
	
			# (set-MailboxFolderPermission -identity "$($user.UserPrincipalName):\Calendar" -User Default -AccessRights LimitedDetails) ;
			
			# Setzen Lizenz
			Try
				{
					Get-MsolUser -Userprincipalname $user.UserPrincipalName | Set-Msoluser -UsageLocation AT
					Set-MsolUserLicense -Userprincipalname $user.UserPrincipalName -AddLicenses $lizenz
					Write-Host "Lizenz $lizenz f�r $User.UserPrincipalName gesetzt " -ForeGroundColor Green 
				}
			Catch
				{
					Write-Host "Fehler bei Lizenzzuweisung f�r $user.UserPrincipalName" -ForeGroundColor Red
				}
			# Setzen Sprache 
			Try 
				{
					Set-MailboxRegionalConfiguration -Identity $user.UserPrincipalName -LocalizeDefaultFolderName:$true -Language $language -DateFormat "dd.MM.yyyy") ;
					Write-Host "Sprache $language fuer $user.UserPrincipalName gesetzt" -ForeGroundColor Green 
				}
			Catch
				{
					Write-Host "Fehler bei Spracheinstellungen fuer User $user.UserPrincipalName" -ForeGroundColor Red
				}
		
			
			# OWA deaktivieren 
			
			Set-CasMailbox $user.UserPrincipalName -OWAEnabled $false
			Write-Host "OWA f�r User $user.UserPrincipalName deaktiviert" -ForeGroundColor Green 
			
			# Setzen Berechtigungen f�r office365oof@domain.com 
			
			Add-MailboxPermission -identity $user.UserPrincipalName -user office365oof@domain.com -AccessRights FullAccess -Automapping:$False
			Write-Host "Berechtigungen f�r User office365oof@meusburger.com auf Mailbox $user.UserPrincipalName gesetzt " -ForeGroundColor Green 
			
		}
		
		
	Foreach ($forward in $forwardimport)
		{
		# Set Mailforwarder
			If ($forward.Forwarder -eq "*@*"){
				Set-Mailbox -Identity $forward.Mailbox -DeliverToMailboxAndForward $true -ForwardingAddress $forward.ForwarderName }		
			
			Else {}
			
		# Set MessageCopyForSentAsEnabled und/oder MessageCopyForSendOnBehalfEnabled
		
			If ($forward.SendAsItemsCopiedTo -eq "SenderandFrom"){
				Set-Mailbox -identity $forward.Mailbox -MessageCopyForSentAsEnabled $true }
				
			Elseif ($_.SendOnBehalfItemsCopiedTo -eq "SenderandFrom"){
				Set-Mailbox -identity $forward.Mailbox -MessageCopyForSendOnBehalfEnabled $true }
				
			Else {}
				
			If ($forward.GrantSendOnBehalf -ne $null) {
				Set-Mailbox -identity $forward.Mailbox -GrantSendOnBehalfTo $forward.GrantSendOnBehalf
				}
			Else {}
		}
		
		
Get-PSSession | Remove-PSSession		

		
		
		