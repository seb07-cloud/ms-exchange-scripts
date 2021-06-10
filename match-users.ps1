Connect-AzureAD
$users = Get-AzureADUser | Where-Object {($_.DirSyncEnabled -eq $null) -and ($_.UserType -eq "Member")}


foreach ($user in $users){

    try {
        if (Get-ADUser -Filter "UserPrincipalName -eq '$($user.UserPrincipalName)'"){
            Write-Host "'$($user.UserPrincipalName)' found in AD" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "'$($user.UserPrincipalName)' not found in AD" -ForegroundColor Red
    }
}


foreach ($user in $users){

    try {
        if (Get-Mailbox -Identity '$($user.UserPrincipalName)'){
            Write-Host "'$($user.UserPrincipalName)' has a Mailbox" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "'$($user.UserPrincipalName)' has no Mailbox" -ForegroundColor Red
    }
}