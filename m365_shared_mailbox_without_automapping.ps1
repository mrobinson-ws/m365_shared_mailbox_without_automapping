#Requires -Module ExchangeOnlineManagement
[cmdletbinding()]
param()


#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}

$allusers = Get-Mailbox -ResultSize Unlimited
$users = $allUsers | Where-Object RecipientTypeDetails -eq UserMailbox | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -Passthru -Title "Please select the user(s) to Share The Mailbox with" | Select-Object -ExpandProperty UserPrincipalName
$sharedmailbox = $allUsers | Where-Object RecipientTypeDetails -eq SharedMailbox | Sort-Object DisplayName | Select-Object -Property Name,Alias,UserPrincipalName | Out-GridView -Title "Please select the mailbox you are adding the user(s) to" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName

Write-Verbose "Removing $user from $sharedmailbox"
foreach ($user in $users) {
    Write-Verbose "Removing Permnission if it exists, ignore errors in yellow saying The Appropriate Access Control Entry Is Already Present"
    Remove-MailboxPermission -Identity $sharedmailbox -User $user -AccessRights FullAccess -Confirm:$false
    Write-Verbose "Adding Full Access"
    Add-MailboxPermission -Identity $sharedmailbox -User $user -AccessRights FullAccess -AutoMapping:$false
    Write-Verbose "Adding Send As Permission"
    Add-RecipientPermission $sharedmailbox -AccessRights SendAs -Trustee $user -Confirm:$false
}
Write-Verbose "$users have been removed (if needed) and re-added to $sharedmailbox without automapping"
