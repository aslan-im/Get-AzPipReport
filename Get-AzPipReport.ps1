#Requires -module Az.Accounts, Az.Network, ImportExcel

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]
    $ReportName = "AzPipReport_$($(Get-Date).ToString('yyyy.MM.dd')).xlsx"
)

<#
.SYNOPSIS
    Script for collecting the PIP information from all subscriptions in an Azure tenant
.DESCRIPTION
    This script will collect the PIP information from all subscriptions in an Azure tenant and export it to an excel file.
    The script will also create a summary sheet with the number of PIPs per subscription.
.NOTES
    Version: 1.0
    Requires the following modules:
    - Az.Accounts
    - Az.Network
    - ImportExcel
    
.LINK
    https://github.com/aslan-im/Get-AzPipReport
.EXAMPLE
    .\Get-AzPipReport.ps1 -ReportName 'MyPipReport.xlsx'
    This will run the script and export the results to a file called MyPipReport.xlsx
.EXAMPLE
    .\Get-AzPipReport.ps1
    This will run the script and export the results to a file called AzPipReport.xlsx
#>

Import-module ImportExcel, Az.Accounts, Az.Network

if ($ReportName.Split('.')[-1] -ne 'xlsx') {
    $ReportName += '.xlsx'
}

# Check Azure Connectivity

$AzConnection = Get-AzContext
if (!$AzConnection) {
    try {
        Connect-AzAccount -ErrorAction Stop
    }
    catch {
        Write-Error "Could not connect to Azure. Please check your credentials and try again."
        exit 1
    }
}

$Subs = get-azSubscription
$AzSubscriptionsCount = $Subs | Measure-Object | Select-Object -ExpandProperty Count
$SubsCounter = 1

$AllPipsResult = @()
$PerSubReport = @()

foreach ($Sub in $Subs) {
    $SubsPercent = $SubsCounter / $AzSubscriptionsCount * 100

    $SubscriptionProgressSplat = @{
        Activity        = "Working on subscription: `"$($Sub.name)`"..."
        PercentComplete = $SubsPercent 
        Status          = "Progress $SubsCounter/$AzSubscriptionsCount ->"
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
        @{L = 'Subscription'; E = { $Sub.name } }
    )
    $Pips = Get-AzPublicIpAddress | Select-Object $PipsSelector

    $PipsCount = $Pips | Measure-Object | Select-Object -ExpandProperty Count
    $PipsCounter = 1

    Write-Output "Found $PipsCount PIPs in $($Sub.name)"
    $PerSubReport += [PSCustomObject]@{
        Subscription = $Sub.name
        PIPsCount    = $PipsCount
    }
    foreach ($Pip in $Pips) {
        $PipsPercent = $PipsCounter / $PipsCount * 100

        $PipsProgressSplat = @{
            Activity        = "Working on PIP: `"$($Pip.name)`"..."
            PercentComplete = $PipsPercent 
            Status          = "Progress $PipsCounter/$PipsCount ->"
            id              = 1
        }
        Write-Progress @PipsProgressSplat

        $AttachedTo = $Pip.IpConfiguration.Id
        $AttachedToType = ''
        $AttachedToName = ''
        if ($AttachedTo) {
            $AttachedToString = $AttachedTo.Split('/')[7]

            if ($AttachedToString -eq 'networkInterfaces') {
                $NICObject = Get-AzNetworkInterface -Name $($AttachedTo.Split('/')[8])
                if ($NICObject.VirtualMachine) {
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
            else {
                $AttachedToType = $AttachedToString
                $AttachedToName = $AttachedTo.Split('/')[8]
            }
        }
        else {
            $AttachedToType = 'Not attached'
            $AttachedToName = 'Not attached'
        }
            

        $AllPipsResultSelector = @(
            "Name",
            "Location",
            "ResourceGroupName",
            "IpAddress",
            "PublicIpAllocationMethod",
            @{L = 'Subscription'; E = { $Sub.name } },
            @{L = 'AttachedToObjectType'; E = { $AttachedToType } },
            @{L = 'AttachedToObjectName'; E = { $AttachedToName } },
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
    Path          = $ReportName
    AutoSize      = $true
    AutoFilter    = $true
    TableStyle    = 'Medium2'
    WorksheetName = 'Summary'
    InputObject   = $PerSubReport
    ErrorAction   = 'Stop'
}

$AllPipsResultSplat = @{
    Path          = $ReportName
    AutoSize      = $true
    AutoFilter    = $true
    TableStyle    = 'Medium2'
    WorksheetName = 'PIPList'
    InputObject   = $AllPipsResult
    ErrorAction   = 'Stop'
}

try{
    Export-Excel @PerSubSplat
    Export-Excel @AllPipsResultSplat
}
catch {
    Write-Error "Could not export to excel. Please check if the file is open and try again."
    exit 1
}
