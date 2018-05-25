Function New-CMWorkstationQuery {
    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'medium')]
    Param (
        [Parameter(Mandatory,
            HelpMessage = 'Please enter the email address you want to send the report from')]
        [ValidateNotNullOrEmpty()]
        [Alias('EmailAddress')]
        [string]$fromEmailAddress,

        [Parameter(Mandatory,
            HelpMessage = 'Please enter your SMTP address password. The traffic is encrypted.')]
        [ValidateNotNullOrEmpty()]
        [Alias('Password')]
        [psobject]$fromEmailPassword,

        [Parameter(Mandatory,
            HelpMessage = 'Please enter your SMTP server. For example: smtp.gmail.com')]
        [ValidateNotNullOrEmpty()]
        [string]$smtpServer,

        [Parameter(Mandatory,
            HelpMessage = 'Please enter your SMTP server port')]
        [ValidateNotNullOrEmpty()]
        [Alias('Port')]
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
        Try {
            $invokeQuery = Invoke-CMWmiQuery -Query $lastBootTimeQuery -Option Lazy
            Foreach ($Query1 in $invokeQuery) {
                $invokeQueryOBJECT = [pscustomobject] @{
                    'HostResults' = $Query1.SMS_R_System
                }
            
                $invokeQueryOBJECT
            }
        
            $queryUser = Invoke-CMWmiQuery -Query $selectUsername -Option Lazy
            $emailUser = $queryUser.LastLogonUserName  | Sort-Object -Unique *
        
            if ($PSCmdlet.ShouldProcess($fromEmailAddress)) {
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
        }

        Catch {
            Write-Output 'Running error test 1/2'
            Pause
            $Lookup = Resolve-DnsName $smtpServer
            if ($Lookup.NameHost) {
                Write-Output 'Connection to SMTP server successful'
                $Lookup
            }

            elseif ($Lookup.Section -notlike "Answer") {
                Write-Output 'Connection unsuccessful'
            }

            Write-Output 'Running error test 2/2'
            $sccmServer = Read-Host 'Please provide the hostname to your SMTP server'
            $testConnection = test-connection $sccmServer
            if ($testConnection.IPV4Address.IPAddressToString[0]) {
                Write-Output 'Connection established'
            }

            elseif (-not($testConnection.IPV4Address.IPAddressToString[0])) {
                Write-Output 'No connection'
            }

            Write-Warning 'An unknown error has occured. Please review..'
            $PSCmdlet.ThrowTerminatingError($_)

        }
    }#Process
    End {}    
}#Function
