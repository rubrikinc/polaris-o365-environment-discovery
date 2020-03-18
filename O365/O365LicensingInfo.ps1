
################################################################################################################################################################
# Script accepts 3 parameters from the command line
#
# Office365Username - Mandatory - Administrator login ID for the tenant we are querying
# Office365Password - Mandatory - Administrator login password for the tenant we are querying
#
#
# To run the script
#
# .\O365LicensingInfo.ps1 -Office365Username admin@xxxxxx.onmicrosoft.com -Office365Password Password123
#
#
# Author: Manoj Verma
# Version: 1.0
################################################################################################################################################################
#Accept input parameters
Param(
[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
[string] $Office365Username,
[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
[string] $Office365Password
)
Function Main 
{
    #Call ConnectTo-ExchangeOnline function with correct credentials
    #ConnectTo-ExchangeOnline -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password
    Connect-MsolService
    Get-MsolAccountSku | FL
}
# Start script
. Main

