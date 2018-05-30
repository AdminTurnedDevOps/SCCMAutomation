Function Get-SCCMDeviceCollection {
    [cmdletbinding()]
    Param()

    $deviceCollections = Get-CMDeviceCollection
    Foreach ($object in $deviceCollections) {
        $deviceCollectionsOBJECT = New-Object PSObject -Property @{
            'Name'            = $Object.Name
            'LastRefreshTime' = $object.LastRefreshTime
            'ID'              = $object.LimitToCollectionID

        }
        $deviceCollectionsOBJECT
    }
    $deviceCollectionsOBJECT | Out-Null

    Foreach ($device in $deviceCollectionsOBJECT) {

        $deviceArray = @()
        $deviceArray += $device
    }

    $deviceArray

}#Function

Function Move-CMDeviceToCollection {
    [cmdletbinding(SupportsShouldProcess = $true, ConfirmImpact = 'high')]
    Param (
        [Parameter(Position = 0,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Please enter the devices hostname you wish to move')]
        [string]$Hostname,

        [Parameter(ParameterSetName = "CollectionName",
        Position = 1,
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Please enter the collection name for the target')]
        [string]$CollectionName,
        
        [Parameter(ParameterSetName = "CollectionID",
        Position = 1,
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Please enter the collection ID for the target')]
        [string]$CollectionID,

        [Parameter(Position = 2,
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Please enter the collection ID rom where the target originated')]
        [string]$deviceOriginatingCollection
    )
    begin {Write-Host -ForegroundColor Red 'If you need a list of device collection names or IDs, please run the other function in this module called Get-SCCMDeviceCollection'}

    process {
        try {
            if ($PSCmdlet.ShouldProcess($CollectionName) -or $PSCmdlet.ShouldProcess($CollectionID)) {
                
                if ($CollectionID -match '\w') {
                    $Option1 = Read-Host "Would you like to move $($Hostname)? 1 for yes or 2 for no"
                    Foreach ($Host1 in $Hostname) {      
                    switch ($Option1) {
                        '1' {
                            $Params1 = @{
                                'CollectionID' = $CollectionID
                                'ResourceId'   = (Get-CMDevice -CollectionName $deviceOriginatingCollection -Name $Host1).ResourceID
                            }
                            Add-CMDeviceCollectionDirectMembershipRule @Params1                   
                        }

                        '2' {
                            Write-Output 'Goodbye'
                            Break
                        }
                    }#Switch
                }
                }#if
                if ($CollectionName -match '\w') {
                    $Option2 = Read-Host "Would you like to move $($Hostname)? 1 for yes or 2 for no"
                    Foreach ($host2 in $Hostname) {
                    switch ($option2) {
                        '1' {
                            $Params2 = @{
                                'CollectionID' = $CollectionName
                                'ResourceId'   = (Get-CMDevice  -CollectionName $deviceOriginatingCollection -Name $host2).ResourceID
                            }
                            Add-CMDeviceCollectionDirectMembershipRule @Params2                
                        }
                        '2' {
                            Write-Output 'Goodbye'
                            Break
                        }
                    }#Switch
                    }
                }#IF
            }#IFPSCMDLET
        }#Try
    
        Catch {
            Write-Warning 'An error has occured while adding your collection name or collection ID'
            $PSCmdlet.ThrowTerminatingError($PSItem)
        }
    }#Process
    end {}
}#Function
