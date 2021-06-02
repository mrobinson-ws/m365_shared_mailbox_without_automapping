#Requires -Modules AzureAD, ExchangeOnlineManagement

# Test And Connect To AzureAD If Needed
try {
    Write-Verbose -Message "Testing connection to Azure AD"
    Get-AzureAdDomain -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Azure AD"
}
catch {
    Write-Verbose -Message "Connecting to Azure AD"
    Connect-AzureAD
}

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

$users = Get-AzureADUser | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -Passthru -Title "Please select the user(s) to Share The Mailbox with" | Select-Object -ExpandProperty UserPrincipalName
$sharedmailbox = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Sort-Object DisplayName | Select-Object -Property Name,Alias,UserPrincipalName | Out-GridView -Title "Please select the mailbox you are adding the user(s) to" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName

foreach ($user in $users) {
    Remove-MailboxPermission -Identity $sharedmailbox -User $user -AccessRights FullAccess -Confirm:$false
    Add-MailboxPermission -Identity $sharedmailbox -User $user -AccessRights FullAccess -AutoMapping:$false
}
Write-Host "$users have been removed and re-added to $sharedmailbox without automapping"
