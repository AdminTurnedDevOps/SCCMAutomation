Function Get-SCCMMissingApp {
    [cmdletbinding(DefaultParameterSetName = 'SCCMQueryMissingApp', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please enter a collection name for your new SCCM collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [ValidateNotNullOrEmpty()]
        [Alias('Name', 'Collection')]
        [psobject]$CollectionName,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please enter the limited collection name for your new device collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [ValidateNotNullOrEmpty()]
        [Alias('LimitingCollection')]
        [psobject]$LimitingCollectionName,

        [Parameter(HelpMessage = 'Please type in a comment/description for your SCCM collection',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [string]$Comment,

        [Parameter(HelpMessage = 'The error log is for any errors that occur. They will be stored in the specified location',
            ParameterSetName = 'SCCMQueryMissingApp')]
        [string]$ErrorLog = (Read-Host 'Please enter a UNC path for an error log if an error occurs in your script')
    )
    Begin {   
        Write-Warning 'Please ensure you have started a PowerShell connection connecting to your appropriate SCCM environment.'
    }

    Process {
        Try {
            Write-Verbose 'Collecting New-CMDeviceCollection parameters'
            $NewCMDeviceCollectionSPLAT = @{
                'Name'                   = $CollectionName
                'LimitingCollectionName' = $LimitingCollectionName
                'RefreshType'            = 'Manual'
                'Comment'                = $Comment
                'Confirm'                = $True
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
            $NEWCMDeviceCollectionPSOBJECT

            IF ($NEWCMDeviceCollection.Name -contains "\w+") {

                Write-Output "The collection $CollectionName has been created. Querying the missing application will now begin"
                $RuleName = $CollectionName += 'MissingApp'
                #Query missing application from device collection
                $AddDeviceCollectionQuerySPLAT = @{
                    'CollectionName'  = $CollectionName
                    'CollectionID'    = $NEWCMDeviceCollection.CollectionID
                    'RuleName'        = $RuleName
                    'QueryExpression' = "select SMS_R_SYSTEM.ResourceID,
                                    SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,
                                    SMS_R_SYSTEM.SMSUniqueIdentifier,
                                    SMS_R_SYSTEM.ResourceDomainORWorkgroup,
                                    SMS_R_SYSTEM.Client from SMS_R_System  inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId  where SMS_G_System_COMPUTER_SYSTEM.Name not in  (select distinct SMS_G_System_COMPUTER_SYSTEM.Name  from SMS_R_System  inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId  inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId  where SMS_G_System_ADD_REMOVE_PROGRAMS.DisplayName like " % $Application % ")"
                }
                $AddDeviceCollectionQuery = Add-CMDeviceCollectionQueryMembershipRule @AddDeviceCollectionQuerySPLAT
    
                Write-Verbose 'Starting query'
                $AddDeviceCollectionQuery
            }#IF

            else {
                Write-Output 'No device collection was created. Please try re-running the script'
            }#ELSE
        }#TRY
        Catch {
            Write-Warning 'An error has occured. Please review the error logs'
            $_ | Out-File $ErrorLog
            #Throw error to host
            Throw
        }
    }#Process 
    End {Write-Verbose 'The function has completed'}
}#Function
