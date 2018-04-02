Function Get-SCCMMissingApp {
    [cmdletbinding(DefaultParameterSetName = 'SCCMQueryMissingApp', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    Param (
        [Parameter(Position = 0,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please enter a collection name for your new SCCM collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [ValidateNotNullOrEmpty()]
        [Alias('Name', 'Collection')]
        [psobject]$CollectionName,

        [Parameter(Position = 1,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please enter an application name for your new SCCM collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [ValidateNotNullOrEmpty()]
        [Alias('LimitingCollection')]
        [psobject]$LimitingCollectionName,

        [Parameter(Position = 2,
            HelpMessage = 'Please type in a comment/description for your SCCM collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [string]$Comment,

        [Parameter(Position = 3,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please type in the appropriate application name that you want to query')]
        [string]$Application = 'UltraVNC',

        [Parameter(Position = 4 ,
            HelpMessage = 'The error log is for any errors that occur. They will be stored in the specified location',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [string]$ErrorLog = (Read-Host 'Please enter a UNC path for an error log if an error occurs in your script')
    )
    Begin {   
        Write-Warning 'Please ensure you have started a PowerShell Window connecting to your appropriate SCCM environment.'
    }

    Process {
        Try {
            Write-Verbose 'Collecting New-CMDeviceCollection parameters'
            $NewCMDeviceCollectionSPLAT = @{
                'Name'                   = $CollectionName
                'LimitingCollectionName' = $LimitingCollectionName
                'RefreshType'            = 'Manual'
                'Comment'                = $Comment
                'Confirm'                = $true
            }
            Write-Verbose 'Creating new device collection'
            $NEWCMDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollectionSPLAT

            Write-Verbose 'Creating new custom object'
            $NEWCMDeviceCollectionPSOBJECT = [pscustomobject] @{
                'CollectionID'         = $NEWCMDeviceCollection.CollectionID
                'SmsProviderObectPath' = $NEWCMDeviceCollection.SmsProviderObectPath
                'LastChangeTime'       = $NEWCMDeviceCollection.LastChangeTime
                'LimitingCollectionTo' = $NEWCMDeviceCollection.LimitingCollectionName
                'CollectionName'       = $NEWCMDeviceCollection.Name
            }

            Write-Verbose 'Outputting custom object based on New-CMDeviceCollection output'
            $NEWCMDeviceCollectionPSOBJECT | ft

            IF ($PSBoundParameters.ContainsKey('Application')) {
                Write-Output "The collection $CollectionName has been created. Querying the missing application will now begin"
                $RuleName = $CollectionName += 'MissingApp'
                #Query missing application from device collection
                $AddDeviceCollectionQueryPARAMS = @{
                    'CollectionID'    = $NEWCMDeviceCollection.CollectionID
                    'RuleName'        = $RuleName
                    'QueryExpression' = "select distinct SMS_R_System.Name, SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName from SMS_R_System inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId inner join SMS_G_System_CH_ClientSummary on SMS_G_System_CH_ClientSummary.ResourceID = SMS_R_System.ResourceId where SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName not like '%$Application%'"
                }
                $AddDeviceCollectionQuery = Add-CMDeviceCollectionQueryMembershipRule @AddDeviceCollectionQueryPARAMS
    
                Write-Verbose 'Starting query'
                $AddDeviceCollectionQuery
            }


        }#TRY
        Catch {
            Write-Warning 'An error has occured. Please review the error logs'
            $_ | Out-File -Encoding utf8 -FilePath ('FileSystem::' + $ErrorLog)
            #Throw error to host
            Throw
        }
    }#Process 
    End {Write-Verbose 'The function has completed'}
}#Function
