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


# Setzen Mailbox Einstellungen

	Foreach ($user in $users)
		{
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
					Set-MailboxRegionalConfiguration -Identity $user.UserPrincipalName -LocalizeDefaultFolderName:$true -Language $language -DateFormat "dd.MM.yyyy")
					Write-Host "Sprache $language fuer $user.UserPrincipalName gesetzt" -ForeGroundColor Green 
				}
			Catch
				{
					Write-Host "Fehler bei Spracheinstellungen fuer User $user.UserPrincipalName" -ForeGroundColor Red
				}
		}