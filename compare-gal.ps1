# compare-GAL
#
# Load Exchange Recipients in Hybrid and compare gal.
# use PrimarySmtpAddress as Matchkey
#
# Pending: Additional Check
#   RecipientTypeDetailscloud=RecipientTypeDetailsOnPrem
#   CloudOnly recipients with maialdresse <> onmicrosoft.com
# 
# 20190917  Fix ProxyAddresses mit SMTPAdressstring
# 20200128  Erweiterung um LegacyDN als X500 Adresse zu addieren
# 20200129  Sonderfall Objekt ohne ProxyAddress abgefangen und LegacyDN-Erweiterung
# 20210526 FC Exchange Remote PowerShell addiert
# 20210527 FC ExchangeOnlineV2 Support, Session reuse support, Bugfixing

param (
    $adminusername = "admin@tenantname.onmicrosoft.com",
    $gallistcsv = ".\gallist.csv", 
    $proxyaddresslistcsv = ".\proxyaddresslist.csv", 
    $resultsize="unlimited",
    [string]$exchangeuri  	= ""
    #[string]$exchangeuri  	= "https://ex01.msxfaq.de/Powershell"
)

Set-PSDebug -Strict

Write-Host "Compare-GAL: Start"

##########################################################
# Loading Exchange environment
##########################################################
Write-Host "Loading Exchange Environment"

#Alt. not supportet https://blog.rmilne.ca/2015/01/28/directly-loading-exchange-2010-or-2013-snapin-is-not-supported/
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
$error.Clear()
if (get-pssession | where-object {$_.State -eq "Opened" -and $_.Runspace.ConnectionInfo.ConnectionUri -eq $exchangeuri}) {
    Write-Host "Using existing Exchange OnPremises Connection"
    $ExOnPremsession = $null
}
else {
    Write-Host "Start new Connection to Exchange OnPremises"
    if ([string]::IsNullOrEmpty($exchangeuri)) {
        Write-host "Retrieving Exchange URL from Active Directory"
        $RootDSE = [adsi]"LDAP://rootDSE"
        $exchangeuri = Get-ADObject `
                            -LDAPFilter "(&(objectclass=msExchPowerShellVirtualDirectory)(msExchInternalHostName=*))" `
                            -SearchBase $RootDSE.configurationNamingContext[0] `
                            -properties msExchInternalHostName `
                        | Select-Object -ExpandProperty msExchInternalHostName -First 1
    }
    
    if ([string]::IsNullOrEmpty($exchangeuri)) {
        Write-Host "No valid Exchange URI found. Check AD or parameters" -ForegroundColor red
        exit
    }
    Write-Host " Connecting to Exchange OnPremises at $($exchangeuri)"
    $ExOnPremSession = New-PSSession `
        -ConfigurationName "Microsoft.Exchange" `
        -ConnectionUri $exchangeuri 
    Write-Host "Import Exchange Remote Session Commandlets"
    import-pssession -Session $ExOnPremSession -AllowClobber | out-null
    set-adserversettings -viewentireforest $true
}
if ($error) {
	Write-Host "Exitcode:8 Unable to load Exchange Snapins"
	exit 8
}
if (get-command "get-recipient" -ErrorAction SilentlyContinue) {
    Write-Host " Commandlet get-recipient found OK"
}
else {
    Write-Host " Unable to load exchange powershell module  Stop"
    exit
}



Write-Host "Connect to Exchange Online Credentials"
if (get-pssession | where-object {$_.State -eq "Opened" -and $_.name -like "ExchangeOnlineInternalSession*"}) {
    Write-Host "Using existing Exchange online Connection"
}
else {
    Write-Host "Start Connection to Exchange Online"
    Connect-ExchangeOnline -ShowBanner:$false
}


$summary = [pscustomobject][ordered]@{
        totalrecipientsonprem=0
        totalrecipientscloud=0
        totaladdressesonprem=0
        totaladdressescloud=0
        totalduplicateproxyonprem=0
        totalerroronprem=0
        totalerrorcloud=0
        mismatch=0
        OnlyInEXO=0
        OnlyOnPrem=0
        ProxyOnlyinCloud=0
        PrimaryMismatch=0
}

[hashtable]$gallist=@{}
[hashtable]$proxyaddresslist=@{}

Write-Host " Preloading Mailboxes to collect LegacyExchangeDN - may take a while"
[hashtable]$LegacyExchangeDNtable=@{}
foreach ($mailbox in get-mailbox -resultsize $resultsize){
    $LegacyExchangeDNtable[$mailbox.primarysmtpaddress]= $mailbox.LegacyExchangeDN
}

Write-Host " Loading OnPremRecipients.. may take a while"
foreach ($recipient in (get-recipient -resultsize $resultsize)) {
    $PrimarySmtpAddress= $recipient.PrimarySmtpAddress.tostring().tolower()
    $summary.totalrecipientsonprem++
    Write-Host "OnPrem:($($summary.totalrecipientsonprem)) $($PrimarySmtpAddress) " -nonewline
    if ($gallist[$PrimarySmtpAddress]){
        Write-Host "  Duplicate Primary SMTP-Address found - SKIP" -BackgroundColor red
        $summary.totalerroronprem++
    }
    else {
        Write-Host "Add" -ForegroundColor white -BackgroundColor blue
        $gallist[$primarysmtpaddress] = [pscustomobject][ordered]@{
            primarysmtpaddress = $PrimarySmtpAddress
            RecipientTypeDetailsOnPrem=$recipient.RecipientTypeDetails.tostring()
            RecipientTypeDetailscloud=""
            status = "OnPrem"
        }
    }

    Write-Host "   Processing EmailAddresses"
    $ProxyAddresses  = ($recipient.emailaddresses)
    if ($LegacyExchangeDNtable[$PrimarySmtpAddress]) {
        Write-Host "    Add LegacyDN $($LegacyExchangeDNtable[$PrimarySmtpAddress])"
        $ProxyAddresses+= "X500:$($LegacyExchangeDNtable[$PrimarySmtpAddress])"
    }

    if ($ProxyAddresses) {
        foreach ($proxyaddress in $ProxyAddresses) {
            if ($proxyaddress.tolower().Startswith("smtp:") -or $proxyaddress.tolower().Startswith("x500:")) {
                Write-Host "    ProxyAddress:$($proxyaddress)" -NoNewline
                $summary.totaladdressesonprem++
                if ($proxyaddresslist[$proxyaddress]){
                    Write-Host " Duplicate " -ForegroundColor red
                    $proxyaddresslist[$proxyaddress].onpremdup=$true
                    $summary.totalduplicateproxyonprem++
                }
                else {
                    Write-Host " Add " -ForegroundColor green
                    $proxyaddresslist[$proxyaddress]= [pscustomobject][ordered]@{
                        proxytyp = $proxyaddress.split(":")[0]
                        proxyaddress = $proxyaddress
                        primarysmtpaddress = $primarysmtpaddress
                        RecipientTypeDetailsOnPrem=$recipient.RecipientTypeDetails.tostring()
                        RecipientTypeDetailscloud=""
                        match=$false
                        onprem=$true
                        EXO=$false
                        onpremdup=$false
                    }
                }
            }
            else {
                #skip non SMTP address
            }
        }
    }
    else {
        write-Warning " Objekt with empty proxyaddresses"
    }
}
Write-Host "Summary totaladdressesonprem $($summary.totaladdressesonprem)"
Write-Host "Summary totalrecipientsonprem $($summary.totalrecipientsonprem)"
Write-Host "Summary totalduplicateproxyonprem $($summary.totalduplicateproxyonprem)"
Write-Host "Summary totalerroronprem $($summary.totalerroronprem)"

Write-Host "Remove Exchange OnPremises Session"
if ($ExOnPremSession) {
    Remove-PSSession $ExOnPremSession
}

#$gallist

Write-Host "Compare-GAL:Connect to Exchange Online Server"

if (!(get-command "get-EXORecipient" -ErrorAction SilentlyContinue)) {
    Write-Host " Commandlet get-recipient found OK - Stopping" -ForegroundColor red
    exit 1
}

Write-Host "Collecting Exchange Online Users"
foreach ($recipient in (Get-EXORecipient -resultsize $resultsize)) {
    $PrimarySmtpAddress= $recipient.PrimarySmtpAddress
    $summary.totalrecipientscloud++
    Write-Host "EXO($($summary.totalrecipientscloud)):$($primarysmtpaddress)" -nonewline
    if ($gallist[$primarysmtpaddress]){
        Write-Host "  SMTPFound " -ForegroundColor green -NoNewline
        $gallist[$primarysmtpaddress].RecipientTypeDetailscloud=$recipient.RecipientTypeDetails.tostring()
        if ($gallist[$primarysmtpaddress].status -eq "OnPrem"){
            Write-Host "  Match " -ForegroundColor green -NoNewline
            $gallist[$primarysmtpaddress].status="MATCH"
        }
        else{
            Write-Host "  Mismatch " -ForegroundColor red -NoNewline
            $gallist[$primarysmtpaddress].status="Mismatch"
            $summary.mismatch++
        }
    }
    else {
        Write-Host " OnlyInEXO" -ForegroundColor yellow
        $summary.OnlyInEXO++
        $gallist[$primarysmtpaddress] = [pscustomobject][ordered]@{
            primarysmtpaddress = $primarysmtpaddress
            RecipientTypeDetailsOnPrem=""
            RecipientTypeDetailscloud=$recipient.RecipientTypeDetails.tostring()
            status = "OnlyinEXO"
        }
    }

    Write-Host " Processing EmailAddresses"
    $ProxyAddresses  = ($recipient.emailaddresses)
    if ($recipient.LegacyExchangeDN) {
        $ProxyAddresses+= "X500:$($recipient.LegacyExchangeDN)"
    }
    if ($ProxyAddresses) {
        foreach ($proxyaddress in $ProxyAddresses) {
            $summary.totaladdressescloud++
            Write-Host "  ProxyAddress:$($proxyaddress)" -NoNewline
            if ($proxyaddress.tolower().Startswith("smtp:") -or $proxyaddress.tolower().Startswith("x500:") ) {
                if ($proxyaddresslist[$proxyaddress]){
                    Write-Host " Found " -ForegroundColor green -nonewline
                    if ($proxyaddresslist[$proxyaddress].primarysmtpaddress = $PrimarySmtpAddress){
                        Write-Host "PrimaryMatch" -ForegroundColor green
                        $proxyaddresslist[$proxyaddress].match=$true
                        $proxyaddresslist[$proxyaddress].RecipientTypeDetailscloud = $recipient.RecipientTypeDetails.tostring()
                    }
                    else {
                        Write-Host "PrimaryMisMatch" -ForegroundColor yellow
                        $summary.PrimaryMismatch++
                    
                    }
                }
                else {
                    Write-Host " Only in Cloud " -ForegroundColor yellow
                    $proxyaddresslist[$proxyaddress] = [pscustomobject][ordered]@{
                        proxytyp = $proxyaddress.split(":")[0]
                        proxyaddress = $proxyaddress
                        primarysmtpaddress = $primarysmtpaddress
                        RecipientTypeDetailsOnPrem =""
                        RecipientTypeDetailscloud = $recipient.RecipientTypeDetails.tostring()
                        match=$false
                        onprem=$false
                        EXO=$true
                        onpremdup=$false
                    }
                    $summary.ProxyOnlyinCloud++
                }
            }
            else {
                Write-Host " Skip nonSMTP"
            }
        }
    }
    else {
        Write-warning " Object without proxy address found"
    }
}
$summary.OnlyOnPrem = (@($gallist.values | Where-object {$_.status -eq "OnPrem"})).count

Write-Host "Writing CSV-Files"
$gallist.Values | export-csv $gallistcsv -NoTypeInformation
$proxyaddresslist.Values | export-csv $proxyaddresslistcsv -NoTypeInformation

Write-Host "Compare-GAL: End"
$summary