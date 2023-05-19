
$Subs = get-azSubscription
$AzSubscriptionsCount = $Subs | Measure-Object | Select-Object -ExpandProperty Count
$SubsCounter = 1

$ReportPath = $PSScriptRoot + "\AzPipReport.csv"

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
            if($($AttachedTo.Split('/')) -contains 'networkInterfaces'){
                $NIC = $AttachedTo.Split('/')[8]
                $NICObject = Get-AzNetworkInterface -Name $NIC
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
            elseif ($($AttachedTo.Split('/')) -contains 'bastionHosts'){
                $Bastion = $AttachedTo.Split('/')[8]
                $AttachedToType = 'Bastion'
                $AttachedToName = $Bastion
            }
            elseif ($($AttachedTo.Split('/')) -contains 'applicationGateways'){
                $AG = $AttachedTo.Split('/')[8]
                $AttachedToType = 'AppGateway'
                $AttachedToName = $AG
            }
            elseif ($($AttachedTo.Split('/')) -contains 'loadBalancers'){
                $LB = $AttachedTo.Split('/')[8]
                $AttachedToType = 'LoadBalancer'
                $AttachedToName = $LB
            }
            elseif ($($AttachedTo.Split('/')) -contains 'virtualNetworkGateways'){
                $VNG = $AttachedTo.Split('/')[8]
                $AttachedToType = 'VNG'
                $AttachedToName = $VNG
            }
            elseif ($($AttachedTo.Split('/')) -contains 'azureFirewalls') {
                $FW = $AttachedTo.Split('/')[8]
                $AttachedToType = 'FW'
                $AttachedToName = $FW
            }
            else{
                $AttachedToType = 'Unknown'
                $AttachedToName = 'Unknown'
            }
        }
        else{
            $AttachedToType = 'Unknown'
            $AttachedToName = 'Unknown'
        }

        $AllPipsResultSelector = @(
            "Name",
            "Location",
            "ResourceGroupName",
            "IpAddress",
            "PublicIpAllocationMethod",
            @{L='Subscription';E={$Sub.name}}
            @{L='AttachedToObjectType';E={$AttachedToType}}
            @{L='AttachedToObjectName';E={$AttachedToName}}
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