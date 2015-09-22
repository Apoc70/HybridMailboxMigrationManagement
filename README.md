# HybridMailboxMigrationManagement
This is the modified Excel spreadsheet to support UTF8 exports/imports.

Use the following PowerShell code to generate an UTF-8 Excel export for use with the updated Excel spreadsheet.

```
$mbx=Get-Mailbox -resultsize unlimited; $mbx | foreach-object {$UPN = $_.UserPrincipalName; $EmailAddress = $_.PrimarySmtpAddress;$OU = $_.OrganizationalUnit; $Type = $_.RecipientTypeDetails; $_ | Get-MailboxStatistics | select @{Name="UPN";expression={$UPN}},@{Name="EmailAddress";expression={$EmailAddress}},@{Name="Type";expression={$Type}},@{Name="OU";expression={$OU}},DisplayName,@{Name="TotalItemSize(MB)";expression={$_.TotalItemSize.Value.ToMB()}},LastLogonTime}|Export-csv .\Mailboxes_Output.csv -notype	-Encoding UTF8	
```

The file itself is copyright (c) by Michael Hall

* TechNet Gallery: https://gallery.technet.microsoft.com/office/Office-365-Hybrid-Mailbox-84519039
* Blog: http://blogs.technet.com/b/mikehall/archive/2013/06/25/office-365-hybrid-mailbox-migration-management.aspx