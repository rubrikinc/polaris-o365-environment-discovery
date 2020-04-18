
################################################################################################################################################################
# Script accepts 3 parameters from the command line
#
# Office365Username - Mandatory - Administrator login ID for the tenant we are querying
# Office365Password - Mandatory - Administrator login password for the tenant we are querying
# UserIDFile - Optional - Path and File name of file full of UserPrincipalNames we want the Mailbox Size for. Seperated by New Line, no header.
#
#
# To run the script
#
# .\Get-AllMailboxSizes.ps1 -Office365Username admin@xxxxxx.onmicrosoft.com -Office365Password Password123 -InputFile c:\Files\InputFile.txt
#
# NOTE: If you do not pass an input file to the script, it will return the sizes of ALL mailboxes in the tenant. Not advisable for tenants with large
# user count (< 3,000)
#
# Author: Manoj Verma
# Version: 1.0
################################################################################################################################################################
#Accept input parameters
Param(
[Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
[string] $Office365Username,
[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]
[string] $Office365Password,
[Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)]
[string] $UserIDFile
)
#Constant Variables
#$OutputFile = "MailboxSizes.csv" #The CSV Output file that is created, change for your purposes
#Main
Function Main 
{
    #Remove all existing Powershell sessions
    Get-PSSession | Remove-PSSession
    #Call ConnectTo-ExchangeOnline function with correct credentials
    $Office365Username ="admin.o365@rdnn14.onmicrosoft.com" 
    $Office365Password="scaledata12#$"
    ConnectTo-ExchangeOnline -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password
    #Prepare Output file with headers
    #Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,NumberOfItems,MailboxSize" -Encoding UTF8
    #Check if we have been passed an input file path
    if ($userIDFile -ne "")
    {
    #We have an input file, read it into memory
    $objUsers = import-csv -Header "UserPrincipalName" $UserIDFile
    }
    else
    {
    #No input file found, gather all mailboxes from Office 365
        $objUsers = get-mailbox  -ResultSize Unlimited | get-mailboxstatistics | Select DisplayName,DeletedItemCount,ItemCount,TotalDeletedItemSize,TotalItemSize,MessageTableTotalSize,MessageTableAvailableSize,AttachmentTableTotalSize,AttachmentTableAvailableSize,OtherTablesTotalSize,OtherTablesAvailableSize,IsEncrypted,LastInteractionTime,LastLoggedOnUserAccount,LastLogoffTime,LastLogonTime,AssociatedItemCount ,IsValid,IsArchiveMailBox,MailBoxType | Export-CSV  MailboxReport.csv
        $objUsersCal = get-mailbox  -ResultSize Unlimited | Get-MailboxFolderStatistics  –FolderScope 'Calendar'  | Select Identity,Name , FolderPath,CreationTime,FolderId,HiddenItemsInFolder,ItemsInFolder,DeletedItemsInFolder,FolderSize | Export-CSV  CalFile.csv
    }
    #$DataPath = "CalendarUsage.csv"
    #$Results =@()
    #$CalendarResults =@()
    #$obj=@()
    #$objmMilboxUsers=get-mailbox
    #write-host $objmMilboxUsers
    #ForEach ($User in $objmMilboxUsers)
    #{
    #    $UserCalender=Get-MailboxFolderStatistics $User.Alias –FolderScope 'Calendar'  | Select Identity,Name , FolderPath,CreationTime,FolderId,HiddenItemsInFolder,ItemsInFolder,DeletedItemsInFolder,FolderSize 
    #    if ($UserCalender.length -eq 1)
    #    {
    #        $CalendarResults += [pscustomobject]$UserCalender
    #    }
    #    else
    #    {
    #        ForEach ($obj in $UserCalender)
    #        {
    #            $CalendarResults += [pscustomobject]$obj
    #        }
    #    }
    #}
    #write-host "Finished"
    #write-host $CalendarResults
    #$CalendarResults | Export-CSV  "CalendarFile.csv"
    #}

    Get-PSSession | Remove-PSSession
}
###############################################################################
#
# Function ConnectTo-ExchangeOnline
#
# PURPOSE
# Connects to Exchange Online Remote PowerShell using the tenant credentials
#
# INPUT
# Tenant Admin username and password.
#
# RETURN
# None.
#
###############################################################################
function ConnectTo-ExchangeOnline
{
Param(
[Parameter(
Mandatory=$true,
Position=0)]
[String]$Office365AdminUsername,
[Parameter(
Mandatory=$true,
Position=1)]
[String]$Office365AdminPassword
)

#Encrypt password for transmission to Office365
$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force
#Build credentials object
$Office365Credentials = New-Object System.Management.Automation.PSCredential $Office365AdminUsername, $SecureOffice365Password
#Create remote Powershell session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Office365credentials -Authentication Basic –AllowRedirection
#Import the session
Import-PSSession $Session -AllowClobber | Out-Null
}
# Start script
. Main



