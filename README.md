# Get-AzPipReport
Script for collecting the Public IP (PIP) information from all subscriptions in an Azure tenant and exporting it to an Excel file.

## Description
This script collects the PIP information from all subscriptions in an Azure tenant using the Azure PowerShell module. It retrieves the PIP details such as name, location, resource group, IP address, allocation method, and the object it is attached to (e.g., virtual machine, network security group).

The script generates two sheets in the Excel report:

- Summary: Provides a summary of the number of PIPs per subscription.
- PIPList: Contains detailed information about each PIP, including the attached object type and name.

## Prerequisites
- Azure PowerShell module (Az.Accounts, Az.Network)
- ImportExcel module

## Usage
```powershell
.\Get-AzPipReport.ps1 [-ReportName <ReportName>]
```
### Parameters
ReportName (Optional): Specifies the name of the Excel report. Default value is 'AzPipReport.xlsx'. If no file extension is provided, '.xlsx' will be appended.

## Examples
```powershell
.\Get-AzPipReport.ps1 -ReportName 'MyPipReport.xlsx'
```
This command runs the script and exports the results to a file called 'MyPipReport.xlsx'.

```powershell
.\Get-AzPipReport.ps1
```
This command runs the script and exports the results to a file called 'AzPipReport.xlsx'.

## Notes
### Version: 1.0  
### GitHub Repository: https://github.com/aslan-im/Get-AzPipReport