clear
$Leavers = import-csv -Delimiter "," -Path "PATH TO CSV HERE"
$Today = get-date -Format "dd/MM/yyyy"


foreach ($Leaver in $leavers){
$fname = $Leaver.Firstname
$lname = $leaver.Surname
$Forward = $Leaver.EmailAccess
$Finish = $leaver.LeavingDate

if($Today -eq $Finish){
$PSUser = '365 Admin User'
$PSPass = '365 Admin Pass'
$PsPassword = ConvertTo-SecureString -AsPlainText $PSPass -Force
$SecureString = $PsPassword
$PSCreds = New-Object System.Management.Automation.PSCredential($PSUser,$PsPassword)
$Email = (Get-ADUser -Identity "$fname$lname" -Properties Mail).Mail
$SAMName = "$fname$Lname"
$Emailname = "$fname $Lname"
$Pword = "Xt5^2jd0lA1!"
$Securepassword = ConvertTo-SecureString $PWord -AsPlainText -Force
$disabled = "Disabled OU Here"

#Functions
Disable-ADAccount -Identity $SAMName 
Set-ADAccountPassword -Identity $SAMName -NewPassword $SecurePassword
Get-ADUser -Identity $SAMName | ForEach-Object { Move-ADObject -Identity $_.ObjectGUID -TargetPath $disabled }
Get-ADUser -Identity $SAMName -Properties MemberOf | ForEach-Object {
  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}
Set-AdUser -Identity $SAMName -Clear ProxyAddresses
Set-AdUser -Identity $SAMName -Clear Mail
$FrwUsr = $Forward
$AutoRep = "Emails to this mailbox are forwarded to $FrwUsr to respond accordingly."

Connect-MsolService -Credential $PSCreds
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $PSCreds -Authentication Basic -AllowRedirection
Import-PSSession $Session

Set-Mailbox -Identity $EmailName -ForwardingSMTPAddress $FRWUsr
Add-MailboxPermission -Identity $emailname -User $FrwUsr -AccessRights FullAccess -InheritanceType All
Set-MailboxAutoReplyConfiguration -Identity $emailname -AutoReplyState Enabled -ExternalMessage $Autorep -InternalMessage $AutoRep
Set-Mailbox -Type Shared -Identity $Emailname
Start-Sleep -Milliseconds 300
(get-MsolUser -UserPrincipalName $email).licenses.AccountSkuId |
foreach{
    Set-MsolUserLicense -UserPrincipalName $email -RemoveLicenses $_
}

Remove-PSSession *
}
}

if(!($Today -eq $Finish)){
exit
}
exit
