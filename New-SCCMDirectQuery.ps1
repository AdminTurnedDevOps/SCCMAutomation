Function New-SCCMDirectQuery {
    [cmdletbinding(DefaultParameterSetName = 'QueryDevices', SupportsShouldProcess = $true, ConfirmImpact = 'medium')]
    Param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Please choose your device collection that you would like to query from',
            ParameterSetName = 'QueryDevices')]
        [ValidateNotNullOrEmpty()]
        [string]$LimitingCollectionName,

        [Parameter(Mandatory = $true,
            HelpMessage = 'Please put in a name for your new SCCM device collection',
            ParameterSetName = 'QueryDevices')]
        [ValidateNotNullOrEmpty()]
        [string]$NewDeviceCollectionName,

        [ValidateNotNullOrEmpty()]
        [parameter(HelpMessage = 'Please enter the value to query. Example anything with cci-bji in the name, you would put cci-bji',
            ParameterSetName = 'QueryDevices')]
        [string]$CollectionQueryValue,

        [Parameter(ParameterSetName = 'QueryDevices')]
        [string]$LogLocation,

        [Parameter(ParameterSetName = 'QueryDevices')]
        [string]$SCCMServer

    )
    Begin {
        $TestSCCMConnection = Test-Connection $SCCMServer

        IF (-Not ($TestSCCMConnection.StatusCode)) {
            Write-Warning 'No connection to the SCCM server'
        }
    }

    Process {
        Try {
            Write-Verbose 'Confirming parameter input'
            IF ($PSBoundParameters.ContainsKey('NewDeviceCollectionName')) {
                # List out LimitingCollection in either a switch or ValidateSet
                $NewCMDeviceCollectionPARAMS = @{
                    'Name'                   = $NewDeviceCollectionName
                    'LimitingCollectionName' = $LimitingCollectionName
                }
                Write-Verbose 'Limiting device collection to specific collection'
                Write-Verbose 'Creating new device collection'

                $NewCMDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollectionPARAMS

                $NewCMDeviceCollectionOBJECT = [pscustomobject] @{
                    'CollectionName'      = $NewCMDeviceCollection.Name
                    'LimitCollectionName' = $NewCMDeviceCollection.LimitToCollectionName
                    'LimitCollectionID'   = $NewCMDeviceCollection.LimitToCollectionID
                }

                Write-Verbose 'Outputting new collection information'
                $NewCMDeviceCollectionOBJECT
    
            }#PSBound_IF

            Write-Verbose 'Checking to confirm device collection contains new collection and collection ID is correct'
            IF ($NewCMDeviceCollection.Name -match "\w" -or $NewCMDeviceCollection.LimitToCollectionID -match "[^02468]") {
            
                Write-Verbose 'Configuring variable to hold resource ID for limited collection'
                $DeviceDirectRulePARAMS = @{
                    'CollectionName' = $LimitingCollectionName
                    'Name'           = "*$CollectionQueryValue*"
                }
                $DeviceDirectRule = Get-CMDevice @DeviceDirectRulePARAMS
                $ResourceID = ($DeviceDirectRule).ResourceID
            
                Write-Verbose 'Adding direct membership rule'
                $CMDeviceDirectMembershipRulePARAMS = @{
                    'CollectionName' = $NewDeviceCollectionName
                    'ResourceId'     = $ResourceID
                }
                Add-CMDeviceCollectionDirectMembershipRule @CMDeviceDirectMembershipRulePARAMS
            }
        }
        Catch {
            Write-Warning "The value in resource ID is $ResourceID"
            Write-Warning 'If there is no value. The resource ID was empty'

            Write-Warning 'Please review the error logs in your specified destination'
            $_ | Out-File $LogLocation
            Throw
        }
    }#Process
    End {}
}#Function
