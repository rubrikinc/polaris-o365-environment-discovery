# Rubrik SharePoint Metrics Collector

## Overview
This repo contains the PowerShell script to get statistics on O365 environment.

Required Modules:
```
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -RequiredVersion 16.0.8029.0
Install-Module -Name SharePointPnPPowerShellOnline -RequiredVersion 3.21.2005.1
```

## Using ths module
$userid = "admine@domain.com"
$password = "MyPass123!"
$securedPassword = ConvertTo-SecureString $password -AsPlainText -Force
$TenantAdminURL = "https://yourtenant-admin.domain.com"
 .\SharePoint\SPOMetrics.ps1 -User $userid -Password $securePassword -AdminURL $TenantAdminURL

This will create 2 csv files: SPMetricsLibs.csv and SPMetricsList.csv
