function Add-O365License {
    [CmdletBinding()]
    param (
        [switch]$Connect
    )
    
    begin {
        if ($Connect) {
            Connect-O365 
        }
    }
    
    process {
        try {
            $users = Get-MsolUser
            $licenses = Get-MsolAccountSku

            $chosenusers = $users | Sort-Object UserPrincipalName | Out-GridView -PassThru -Title "Choose User(s) you want to assign a license"
            $chosenlicenses = $licenses | Sort-Object AccountSkuId | Out-GridView -PassThru -Title "Choose the license(s) you want to assign" 

            $UsersExistingLicenses = @()
            
            foreach ($user in $chosenusers) {
                $u = get-msoluser -UserPrincipalName $user.UserPrincipalName | Select-Object UserPrincipalName, DisplayName, Licenses, FirstName, LastName 
                
                $UsersExistingLicense = [PSCustomObject]@{
                    UserName       = $u.DisplayName
                }

                for ($z=0; $z -lt $u.Licenses.AccountSkuId.Length; $z++ ){
                    $UsersExistingLicense | Add-Member -type NoteProperty -Name "Lizenz$($z)" -Value $u.Licenses.AccountSkuId[$z]
                }

                $UsersExistingLicenses += $UsersExistingLicense
            }

            $UsersExistingLicenses | Out-GridView -Title "These License are currently assigned to the Users"

            $i = 0
            foreach ($user in $chosenusers){
                Write-Progress -Activity "Processing User" -Id 1 -Status "Processing $i/$($chosenusers.count) User(s)" -PercentComplete ($i / $chosenusers.count * 100)
                foreach ($license in $chosenlicenses) {
                    $t = $t + 1
                    Write-Progress -Activity "Assigning License(s)" -Id 2 -Status "Processing $t/$($chosenlicenses.count) license(s)" -PercentComplete ($t / $chosenlicenses.count * 100)
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $license.AccountSkuId
                }
            }
        }
        catch {
            Write-Host "Could assign license" $license.AccountSkuId "to" $user.UserPrincipalName
        }
    }
    
    end {}
}
