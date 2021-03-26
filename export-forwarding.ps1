$ErrorActionPreference = "SilentlyContinue"

$users = Import-Csv ".\Migrationbatch.csv"

$ForwardExport = @()

# Mailforwarding in PSObject

Foreach ($user in $users) {
    $myexupn = $user.UserPrincipalName
    $myexmailbox = Get-Mailbox $myexupn
    $myexsendonbehalf = $myexmailbox.GrantSendOnBehalfTo
    $myexsendonbehalfaddress = (Get-Mailbox -Identity $myexsendonbehalf.name).PrimarySmtpAddress.ToString()
    $myexforwardingaddress = (Get-Mailbox $myexmailbox.ForwardingAddress).PrimarySmtpAddress.ToString()
    $myexforwardingname = (Get-Mailbox $myexmailbox.ForwardingAddress).Name.ToString()
    $myexitemsconf = Get-MailboxSentItemsConfiguration $myexupn
           
    $ForwardObject = New-Object PSObject -Property @{    
           
        Mailbox                   = $myexupn
        Forwarder                 = $myexforwardingaddress
        ForwarderName             = $myexforwardingname
        SendAsItemsCopiedTo       = $myexitemsconf.SendAsItemsCopiedTo
        SendOnBehalfItemsCopiedTo = $myexitemsconf.SendOnBehalfOfItemsCopiedTo
        GrantSendOnBehalf         = $myexsendonbehalfaddress
    }
    $ForwardExport += $ForwardObject
                     
    If ($myexforwardingaddress -like "*@*") {
        Write-Host "Mailforwarder $myexforwardingaddress auf $myexupn in CSV exportiert" -ForegroundColor Green 
    } 
					
    Else {}
                                       
    Clear-Variable myex*
                                                            
}
                                            
# Mailforwarding in CSV Exportieren                     
                               
$ForwardExport | Export-Csv ".\Forwarding1.csv" -NoTypeInformation
