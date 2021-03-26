function export-immutable {
    [CmdletBinding()]
    param (
        [string]$OU,
        [string]$Group,
        [Parameter(Mandatory, HelpMessage = "Input the Fullpath to the CSV File, include the File Extension!")]
        [System.IO.Path]$path
    )

    begin {
        $CSVExport = @()
    }
    process {
        if ($OU -ne " ") {
            foreach ($user in (Get-Aduser -SearchBase $OU -Filter *)) {
                $guid = (Get-ADUser $user).Objectguid
                $immutableID = [system.convert]::ToBase64String($guid.tobytearray())
            
                $mymail = (Get-Aduser $user -Properties 'Mail').mail
                $myupn = $user.Userprincipalname
                $myimmu = $immutableID
            
                $CSVUser = New-Object PSObject -Property @{  
            
                    Userprincipalname = $myupn
                    Emailaddress      = $mymail
                    ImmutableID       = $myimmu
                }
                $CSVExport += $CSVUser
                Clear-Variable myex*
            }
            $CSVExport | Export-Csv -Path $path-NoTypeInformation
        }
        elseif ($Group -ne " ") {
            foreach ($user in (Get-AdGroupMember -identity )) {

                $guid = (Get-ADUser $user).Objectguid
                $immutableID = [system.convert]::ToBase64String($guid.tobytearray())
            
                $mymail = (Get-Aduser $user -Properties 'Mail').mail
                $myupn = $user.Userprincipalname
                $myimmu = $immutableID
            
            
                $CSVUser = New-Object PSObject -Property @{  
            
                    Userprincipalname = $myupn
                    Emailaddress      = $mymail
                    ImmutableID       = $myimmu
                }
                $CSVExport += $CSVUser
                Clear-Variable myex*
        
                $CSVExport | Export-Csv -Path $path-NoTypeInformation
            }
        }
        else {
            Write-Host "Define a Parameter a Group or Organisation Unit is mandatory" -ForegroundColor Red
        }
    }
    end {}
}


function set-immutable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, HelpMessage = "Input the Fullpath to the CSV File, include the File Extension!")]
        [System.IO.Path]$path
    )
    begin {}
    process {
        foreach ($user in (Import-Csv -Path $path -Delimiter ",")) {
            try {
                Set-MsolUser -UserPrincipalName $user.Emailaddress -ImmutableId $user.ImmutableID 
                Write-Host $immutableID "auf User" $user.Emailaddress "gesetzt" -ForegroundColor Green
            }
            catch {
                Write-Host $_
            }
        }
    }
    end {}
}