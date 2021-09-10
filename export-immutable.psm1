<#
.SYNOPSIS
    Exports and Imports the Immutable ID on AD Objects
    Exports the Immutable ID to a .csv File
    and Imports it on CloudOnly Azure AD Users
.DESCRIPTION
    Exports and Imports the Immutable ID on AD Objects
        Exports the Immutable ID to a .csv File
    and Imports it on CloudOnly Azure AD Users
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\Export-Immutable -OU "OU=Contoso-Groups,DC=Contoso,DC=local" -Path C:\temp\export.csv
    .\Set-Immutable -Path C:\temp\export.csv
.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   30.03.2021

    Changelog:
		1.0             Initial Release
#>

function Export-Immutable {
    [CmdletBinding()]
    param (
        [string]$OU,
        [string]$Group,
        [Parameter(Mandatory, HelpMessage = "Input the Fullpath to the CSV File, include the File Extension!")]
        [String]$path
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
            foreach ($user in (Get-AdGroupMember -identity $Group)) {

                $guid = (Get-ADUser $user).Objectguid
                $immutableID = [system.convert]::ToBase64String($guid.tobytearray())
            
                $mymail = (Get-Aduser $user -Properties 'Mail').mail
                $myupn = $user.Userprincipalname
                $myimmu = $immutableID
            
            
                $CSVUser = New-Object -TypeName PSCustomObject -Property @{
            
                    Userprincipalname = $myupn
                    Emailaddress      = $mymail
                    ImmutableID       = $myimmu
                }
                $CSVExport += $CSVUser
                Clear-Variable myex*
        
                $CSVExport | Export-Csv -Path $path -NoTypeInformation
            }
        }
        else {
            Write-Host "Define a Parameter a Group or Organisation Unit is mandatory" -ForegroundColor Red
        }
    }
    end {}
}


function Set-Immutable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, HelpMessage = "Input the Fullpath to the CSV File, include the File Extension!")]
        [string]$path
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