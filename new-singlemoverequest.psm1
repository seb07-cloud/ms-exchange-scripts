<#
.SYNOPSIS
    Add New Single Moverequest
.DESCRIPTION
    Add new Single Moverequest
.EXAMPLE
    .\New-SingleMoveRequest -User Test@domain.com -RoutingDomain customer.mail.onmicrosoft.com
.INPUTS
    -User Test@domain.com 
    -RoutingDomain customer.mail.onmicrosoft.com
.OUTPUTS
    <None>
.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
	Date : 			   29.03.2021
       
    Changelog:
		1.0             Initial Release
#>

function New-SingleMoveRequest {
    [CmdletBinding(HelpMessage = "Please Provide the Users UserPrincipalName and the Office 365 Routing Domain, ie: <<Customer>.mail.onmicrosoft.com> ")]
    param (
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [string]$RoutingDomain
    )
    begin {

        Import-Module ActiveDirectory, MSOnline, CredentialManager
        # Get Credentials
        $cred = Get-Credential -Message "Please Provide the Tenant Admin Credentials"
        $opcred = Get-Credential -Message "Please Provide the OnPremise Domain Admin Credentials"

        Write-Host "Connecting to Office 365 ......" -ForegroundColor Green

        Connect-MsolService -Credential $cred
        $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
        Import-PSSession $s -AllowClobber

        Write-Host "Getting Office 365 Migration Endpoint ......" -ForegroundColor Green
        $endpoint = Get-MigrationEndPoint

    }
    
    process {
        {
            Try {
                New-MoveRequest -Erroraction Stop -Identity $User -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain $RoutingDomain -RemoteCredential $opcred -SuspendWhenReadyToComplete:$true
                Write-Host 'MoveRequest für ' $User ' erstellt' -ForeGroundColor Green
            }
            Catch {
                Write-Host 'Fehler bei ' $User -ForeGroundColor Red
            }
        }	
                
    }
    
    end {
        Remove-PSSession $s -Confirm:$False
    }
}

