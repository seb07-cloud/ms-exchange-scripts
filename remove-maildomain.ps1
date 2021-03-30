$domain = "@raetikon-treuhand.at"

$query = "*" + $domain

$users = get-mailbox | Where-Object { $_.emailaddresses -like $query }
foreach ($user in $users) {
    $addresses = (get-mailbox $user.alias).emailaddresses
    $fixedaddresses = $addresses | Where-Object  { $_.proxyaddressstring -notlike $query }
    set-mailbox $user.alias -emailaddresses $fixedaddresses
    Write-Host "Removed Maildomain from" $user.Name -ForeGroundColor Green
}