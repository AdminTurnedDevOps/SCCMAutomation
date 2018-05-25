Function New-CMWorkstationQuery {
    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'medium')]
    Param (
        [string]$fromEmailAddress,

        [psobject]$fromEmailPassword,

        [string]$smtpServer,

        [int]$smtpPort

    )

    Begin {    
        Add-Type -AssemblyName Microsoft.ConfigurationManagement.PowerShell.Provider.SmsProviderSearch
        Add-Type -AssemblyName Microsoft.ConfigurationManagement.PowerShell.Provider.SearchParameterOrderBy

        #Query for machines that are pending reboot
        $lastBootTimeQuery = @"
                            select SMS_R_System.Name, SMS_R_System.LastLogonUserName, SMS_G_System_OPERATING_SYSTEM.LastBootUpTime, SMS_G_System_WORKSTATION_STATUS.LastHardwareScan, SMS_G_System_COMPUTER_SYSTEM.SystemType from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_WORKSTATION_STATUS on SMS_G_System_WORKSTATION_STATUS.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId order by SMS_G_System_OPERATING_SYSTEM.LastBootUpTime
"@

        #Query for usernames of the pending reboot machines
        $selectUsername = @"
                        select LastLogonUserName from SMS_R_System
"@
    }

    Process {

        $invokeQuery = Invoke-CMWmiQuery -Query $lastBootTimeQuery -Option Lazy
        Foreach ($Query1 in $invokeQuery) {

            $invokeQueryOBJECT = [pscustomobject] @{
                'HostResults' = $Query1.SMS_R_System
            }
            
            $invokeQueryOBJECT
        }
        
        $queryUser = Invoke-CMWmiQuery -Query $selectUsername -Option Lazy
        $emailUser = $queryUser.LastLogonUserName  | Sort-Object -Unique *
        
        if($PSCmdlet.ShouldProcess($fromEmailAddress)) {
        Foreach ($Person in $emailUser) {
            $emailDomain = Read-Host ('Please enter your email domain. Example: @gmail.com')
            $To = $Person + $emailDomain
            $From = $fromEmailAddress
            $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, ($fromEmailPassword | ConvertTo-SecureString -AsPlainText -Force)
            $SMTPServer = $smtpServer
            $Port = $smtpPort
            $sendMailMessagePARAMS = @{
                'From'       = $From
                'To'         = $To
                'Subject'    = "Pending Machine Reboot: IMPORTANT"
                'Body'       = "Reboots are pending on your machine for security updates. Please reboot by end of $(Get-Date -Format MM.dd.yyyy)."
                'Credential' = $Creds
                'SmtpServer' = $SMTPServer
                'Port'       = $Port
                'UseSsl'     = $true
            }
            Send-MailMessage @sendMailMessagePARAMS
        }#Foreach
        }#If
    }#Process
    End {}
}#Function
