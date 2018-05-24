Add-Type -AssemblyName Microsoft.ConfigurationManagement.PowerShell.Provider.SmsProviderSearch
Add-Type -AssemblyName Microsoft.ConfigurationManagement.PowerShell.Provider.SearchParameterOrderBy

$lastBootTimeQuery = @"
select SMS_R_System.Name, SMS_R_System.LastLogonUserName, SMS_G_System_OPERATING_SYSTEM.LastBootUpTime, SMS_G_System_WORKSTATION_STATUS.LastHardwareScan, SMS_G_System_COMPUTER_SYSTEM.SystemType from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_WORKSTATION_STATUS on SMS_G_System_WORKSTATION_STATUS.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId order by SMS_G_System_OPERATING_SYSTEM.LastBootUpTime
"@

$selectUsername = @"
select LastLogonUserName from SMS_R_System
"@

$invokeQuery = Invoke-CMWmiQuery -Query $lastBootTimeQuery -Option Lazy
Foreach ($Query1 in $invokeQuery) {

    $invokeQueryOBJECT = [pscustomobject] @{
        'HostInformation' = $Query1.SMS_R_System
    }
                         
    $invokeQueryOBJECT | fl
}
                                                                     
$queryUser = Invoke-CMWmiQuery -Query $selectUsername -Option Lazy
$emailUser = $queryUser.LastLogonUserName  | Sort-Object -Unique *
Foreach ($Person in $emailUser) {

    $To = $Person + "@youremaildomain.com"
    $From = 'yoursmtpemail@email.com'
    $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, ('password' | ConvertTo-SecureString -AsPlainText -Force)
    $SMTPServer = 'smtpserver.server.com'
    $Port = 'portnumber'
    $sendMailMessagePARAMS = @{
        'From'       = $From
        'To'         = $To
        'Subject'    = "Pending Machine Reboot: IMPORTANT"
        'Body'       = 'Reboots are pending on your machine for security updates. Please reboot by end of date.'
        'Credential' = $Creds
        'SmtpServer' = $SMTPServer
        'Port'       = $Port
        'UseSsl'     = $true
    }
    Send-MailMessage @sendMailMessagePARAMS
}
