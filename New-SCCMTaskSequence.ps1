Function New-SCCMTaskSequence {

    [Cmdletbinding(DefaultParameterSetName = 'SCCMTaskSequence', SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    Param(

        [Parameter(ParameterSetName = 'SCCMTaskSequence', Position = 0, 
            Mandatory = $true,
            ValueFromPipeline = $true)]
        [Alias('Name')]
        [string]$TSName,

        [Parameter(Position = 1, ParameterSetName = 'SCCMTaskSequence')]
        [Alias('Description')]
        [string]$TSDescription,

        [Parameter(Position = 2, ParameterSetName = 'SCCMTaskSequence')]
        [ValidateSet('PackageIDNumber')]
        [string]$BootImagePackageId,

        [Parameter(Position = 3, 
            ParameterSetName = 'SCCMTaskSequence', 
            HelpMessage = 'Please enter the Package ID which is located under the Package ID column on the task sequence')]
        [string]$OSPackageID,

        [Parameter(Position = 4, 
            ParameterSetName = 'SCCMTaskSequence', 
            HelpMessage = 'Please enter the image index. You can find this in the Task Sequence Editor under Apply operating System')]
        [string]$OSImageIndex,

        [Parameter(Position = 5, 
            ParameterSetName = 'SCCMTaskSequence', HelpMessage = 'Please enter the image name WITHOUT the .WIM')]
        [string]$ImageName,

        [Parameter(Position = 6, 
            ParameterSetName = 'SCCMTaskSequence')]
        [string]$DomainName,

        [Parameter(Position = 7, 
            ParameterSetName = 'SCCMTaskSequence')]
        [string]$DomainAccount,

        [Parameter(Position = 8, ParameterSetName = 'SCCMTaskSequence')]
        [string]$OSFilePath,

        [Parameter(Position = 9, 
            ParameterSetName = 'SCCMTaskSequence')]
        [string]$ErrorLog = "C:\Users\$ENV:USERNAME\Desktop\SCCMTSError.txt"

    )

    Begin {
        #Write your SCCM Host IP below within the single quotes
        $IPAddress = 'Your SCCM Host IP'
        #Write your SCCM host hostname below within the single quotes
        $TestSCCMServer = Test-Connection 'Your SCCM hostname'
        IF ($TestSCCMServer.IPV4Address.IPAddressToString[0] -like $IPAddress) {
            Write-Output 'Connection to server: Successful'
            Pause
        }

        ELSE {
            Write-Warning 'Connection to server: Unsuccessful'
            Pause
            Exit
        }
    }
    
    Process {
        Write-Verbose "Asking for variables to plug into the script"
        #Please enter your password that you'd like to pass in below
        $ADPassword = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ('SCCMADAccountPassword' | ConvertTo-SecureString)

        TRY {
            Write-Verbose "Trying to create task sequence"

            IF ($PSCmdlet.ShouldProcess($TSName) -and $PSBoundParameters.ContainsKey('OSPackageID')) {

                $NewCMTaskSequencePARAMS = @{
                    'Name'                          = $TSName
                    'Description'                   = $TSDescription
                    'BootImagePackageId'            = $BootImagePackageId
                    'OperatingSystemImagePackageId' = $OSPackageID
                    'OperatingSystemImageIndex'     = $OSImageIndex
                    'JoinDomain'                    = 'DomainType'
                    'DomainName'                    = $DomainName
                    'DomainAccount'                 = $DomainAccount
                    'DomainPassword'                = $ADPassword
                    'OperatingSystemFilePath'       = $OSFilePath
                }

                New-CMTaskSequence @NewCMTaskSequencePARAMS | Enable-CMTaskSequence               
            }#IF
        }
        CATCH {
            Write-Output 'Testing task sequence...'
            $GetTaskSequence = Get-CMTaskSequence -Name $TSName
            IF ($GetTaskSequence.Name -contains $TSName) {

                Write-Output 'Task Sequence: Exists. Please check error logs for additional information'
            }

            ELSE {

                Write-Warning 'Task seuqnce: DOES NOT EXIST. Please check the error logs.'
            }

            $_ | Out-File $ErrorLog
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

End {}
}
