<#
.SYNOPSIS
    Script to set calendar permissions
.DESCRIPTION
    Script was designed to set calendar permissions based on which OU the user is located in		
.OUTPUTS
    Nothing but magic
.EXAMPLE
    .\Set-CalendarPermissions.ps1
.NOTES
    Author:            Sebastian Wild	
    Email: 			   sebastian.wild@dynabcs.at
    Company:           DynaBCS Informatik
       
    Changelog:
       1.0             Initial Release
#>


Import-Module ActiveDirectory
$ErrorActionPreference = "SilentlyContinue"

$USER_KUB = "OU=KUGES_User_KUB,DC=domain,DC=local"
$USER_VLM = "OU=KUGES_User_VLM,DC=domain,DC=local"
$USER_VLT = "OU=KUGES_User_VLT,DC=domain,DC=local"
$USER_ZD = "OU=KUGES_User_ZD,DC=domain,DC=local"
$USER_STUSAG = "OU=KUGES_User_STUSAG,DC=domain,DC=local"

$User = $null

#Prompt for AD USer 
while ($null -eq $User){
    [ValidateLength(4,7)]$User = [string](Read-Host -Prompt "Input AD-User (SamAccountName)")
}

$Calendar = $User + ":\kalender"

#Get OU in which the User is located
$OU = Get-ADUser -identity $user -Properties canonicalName | Select-Object -Property canonicalName,DistinguishedName,@{Name='OU';Expression={$_.DistinguishedName.Split(',')[1..$($_.DistinguishedName.Split(',')).count] -join ','}}

#Set Permission Func
function Set-Perm {

Param
    (
	    $inputperm,
        [string]$inputcal
    )

        $cal = $Calendar.ToString()
        foreach ($Permission in $Permissions)
            {
                #$group = $Permission.Group.ToString()
                #$right = $Permission.Perm.ToString()
                Add-MailboxFolderPermission -Identity $cal -User $Permission.Group.ToString() -AccessRights $Permission.Perm.ToString()
            }
}

#which OU ?
switch ($OU.OU) {
    $USER_KUB {
        $Permissions = @()
        $Groups = New-Object [PSCustomObject]@{
            Group = 'KUB_Alle'
            Perm = 'Reviewer'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'ZD_Alle'
            Perm = 'Editor'
            }
        Set-Perm -inputperm $Permissions -inputcal $Calendar
        break
    }
    $USER_VLM {
        $Permissions = @()
        $Permissions += [PSCustomObject]@{
            Group = 'VM_Alle'
            Perm = 'Reviewer'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'VM_Direktion'
            Perm = 'Editor'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'ZD_Alle'
            Perm = 'Editor'
            }
        Set-Perm -inputperm $Permissions -inputcal $Calendar
        break
    }
    $USER_VLT {
        $Permissions = @()
        $Permissions += [PSCustomObject]@{
            Group = 'VLT_Alle'
            Perm = 'Reviewer'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'VLT_Direktion'
            Perm = 'Reviewer'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'ZD_Alle'
            Perm = 'Editor'
            }
        Set-Perm -inputperm $Permissions -inputcal $Calendar
        break
    }
    $USER_ZD {
        $Permissions = @()
        $Permissions += [PSCustomObject]@{
            Group = 'ZD_Alle'
            Perm = 'Editor'
            }
        Set-Perm -inputperm $Permissions -inputcal $Calendar
        break    
    }
    $USER_STUSAG {
        $Permissions = @()
        $Permissions += [PSCustomObject]@{
            Group = 'VM_Alle'
            Perm = 'Reviewer'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'VM_Direktion'
            Perm = 'Editor'
            }
        $Permissions += [PSCustomObject]@{
            Group = 'ZD_Alle'
            Perm = 'Editor'
             }
        Set-Perm -inputperm $Permissions -inputcal $Calendar
        break
    }
}