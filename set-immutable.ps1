foreach ($ADUser in (get-adgroupmember -identity G_O365))
	{
		$365User = $aduser.samaccountname + "@biberkopf.at"
		$guid =(Get-ADUser $ADUser).Objectguid
		$immutableID=[system.convert]::ToBase64String($guid.tobytearray())
        #$365Acc = Get-MsolUser -UserPrincipalName $O365User
		Set-MsolUser -UserPrincipalName $365User -ImmutableId $immutableID
        Write-Host $immutableID "auf User" $ADUser.Name "gesetzt" -ForegroundColor Green
	}