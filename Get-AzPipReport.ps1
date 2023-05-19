#Requires -module Az.Accounts, Az.Network, ImportExcel

Import-module ImportExcel, Az.Accounts, Az.Network
$Subs = get-azSubscription
$AzSubscriptionsCount = $Subs | Measure-Object | Select-Object -ExpandProperty Count
$SubsCounter = 1

$ReportPath = $PSScriptRoot + "\AzPipReport.xlsx"

$AllPipsResult = @()
$PerSubReport = @()

foreach ($Sub in $Subs){
    $SubsPercent = $SubsCounter / $AzSubscriptionsCount * 100

    $SubscriptionProgressSplat = @{
        Activity = "Working on subscription: `"$($Sub.name)`"..."
        PercentComplete = $SubsPercent 
        Status = "Progress $SubsCounter/$AzSubscriptionsCount ->"
    }

    Write-Progress @SubscriptionProgressSplat

    Select-AzSubscription $Sub.name | Out-Null
    Write-Output "Working with subscription $($Sub.name)"
    $PipsSelector = @(
        "Name",
        "Location",
        "ResourceGroupName",
        "IpAddress",
        "IpConfiguration",
        "Id",
        "PublicIpAllocationMethod"
        @{L='Subscription';E={$Sub.name}}
    )
    $Pips = Get-AzPublicIpAddress | Select-Object $PipsSelector

    $PipsCount = $Pips | Measure-Object | Select-Object -ExpandProperty Count
    $PipsCounter = 1

    Write-Output "Found $PipsCount PIPs in $($Sub.name)"
    $PerSubReport += [PSCustomObject]@{
        Subscription = $Sub.name
        PIPsCount = $PipsCount
    }
    foreach ($Pip in $Pips){
        $PipsPercent = $PipsCounter / $PipsCount * 100

        $PipsProgressSplat = @{
            Activity = "Working on PIP: `"$($Pip.name)`"..."
            PercentComplete = $PipsPercent 
            Status = "Progress $PipsCounter/$PipsCount ->"
            id = 1
        }
        Write-Progress @PipsProgressSplat

        $AttachedTo = $Pip.IpConfiguration.Id
        $AttachedToType = ''
        $AttachedToName = ''
        if($AttachedTo){
            $AttachedToString = $AttachedTo.Split('/')[7]
            switch ($AttachedToString) {
                "networkInterfaces" { 
                    $NICObject = Get-AzNetworkInterface -Name $($AttachedTo.Split('/')[8])
                    if($NICObject.VirtualMachine){
                        $VMName = $NICObject.VirtualMachine.Id.Split('/')[8]
                        $AttachedToType = 'VM'
                        $AttachedToName = $VMName
                    }
                    elseif ($NicObject.NetworkSecurityGroup) {
                        $NSG = $NICObject.NetworkSecurityGroup.Id.Split('/')[8]
                        $AttachedToType = 'NSG'
                        $AttachedToName = $NSG
                    }
                }
                "bastionHosts" {
                    $AttachedToType = 'Bastion'
                    $AttachedToName = $AttachedTo.Split('/')[8]
                }
                "applicationGateways" {
                    $AttachedToType = 'AppGateway'
                    $AttachedToName = $AttachedTo.Split('/')[8]
                }
                "loadBalancers" {
                    $AttachedToType = 'LoadBalancer'
                    $AttachedToName = $AttachedTo.Split('/')[8]
                }
                "virtualNetworkGateways" {
                    $AttachedToType = 'VNG'
                    $AttachedToName = $AttachedTo.Split('/')[8]
                }
                "azureFirewalls" {
                    $AttachedToType = 'FW'
                    $AttachedToName = $AttachedTo.Split('/')[8]
                }

                Default {
                    $AttachedToType = 'Unknown'
                    $AttachedToName = "Unknown"
                }
            }
        }
        else{
            $AttachedToType = 'Not attached'
            $AttachedToName = 'Not attached'
        }
            

        $AllPipsResultSelector = @(
            "Name",
            "Location",
            "ResourceGroupName",
            "IpAddress",
            "PublicIpAllocationMethod",
            @{L='Subscription';E={$Sub.name}},
            @{L='AttachedToObjectType';E={$AttachedToType}},
            @{L='AttachedToObjectName';E={$AttachedToName}},
            "Id"
        )

        $AllPipsResult += $Pip | Select-Object $AllPipsResultSelector
        $Pips = $null

        $PipsCounter += 1
    }
    $SubsCounter += 1
}
#Export to excel
$PerSubSplat = @{
    Path = $ReportPath
    AutoSize = $true
    AutoFilter = $true
    TableStyle = 'Medium2'
    WorksheetName = 'Summary'
}

$AllPipsResultSplat = @{
    Path = $ReportPath
    AutoSize = $true
    AutoFilter = $true
    TableStyle = 'Medium2'
    WorksheetName = 'PIPList'
}

$PerSubReport | Export-Excel @PerSubSplat
$AllPipsResult | Export-Excel @AllPipsResultSplat