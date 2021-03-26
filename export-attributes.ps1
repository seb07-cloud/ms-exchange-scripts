Import-Module ActiveDirectory

$date = ([datetime]::UtcNow).tostring("MM_dd")
$filename = "teams_phonenumbers_" + $date + ".csv"

$CSVExport = @()

foreach ($user in (Get-ADGroupMember -Identity "DYNA_Teams-PhoneUsers")){

    $ADUser = Get-ADUser -Identity $user -Properties * 

    if ($ADUser.telephoneNumber){

        $myemailaddress = $ADUser.EmailAddress
        $mytelephonenumber = $ADUser.telephoneNumber

        $CsvUsers = New-Object PSObject -Property @{

        "eMail" = $myemailaddress
        "phoneNumber" = $mytelephonenumber
		
        } 

        $CSVExport += $CsvUsers
        Clear-Variable my*

        }
    else {
        $date = ([datetime]::UtcNow).tostring("MM_dd")
        $logfile = "usernoid" + "_" + $date + ".csv"
        $content = "User" + $AzAdUser.GivenName + "." + $AzAdUser.Surname + "has no phoneNumber, and is excluded from provisioning."
        $content | Out-File -FilePath C:\pccfg\ADExport\error\$logfile -Append
        }
    }

$CSVExport |select-object "eMail","phoneNumber" | Export-CSV "C:\pccfg\ADExport\success\$filename" -NoTypeInformation -Encoding UTF8