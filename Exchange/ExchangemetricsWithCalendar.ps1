
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
[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
[string] $Office365Username,
[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
[Security.SecureString] $Office365Password,
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
    ConnectTo-ExchangeOnline -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password
    #Prepare Output file with headers

    $MailBoxDataPath = "MailboxReport.csv"
    $CalendarDataPath = "CalendarReport.csv"

    #Check if we have been passed an input file path
    if ($userIDFile -ne "")
    {
        #We have an input file, read it into memory
        $objUsers = import-csv -Header "UserPrincipalName" $UserIDFile
        #write-host "*******************"
        #write-host $objUsers
        #write-host "*******************"

        $Results =@()
        $CalendarResults =@()
        $MailboxResults =@()

        $obj=@()
        ForEach ($User in $objUsers)
        {
            
            $UserCalender=Get-MailboxFolderStatistics -Identity $User.UserPrincipalName  –FolderScope 'Calendar'  | 
            Select-object   @{
                            Label = "Identity"
                            Expression = { if ($_.Identity) { $_.Identity } else { "No Data" } }
                            },
                            @{
                            Label = "Date"
                            Expression = { if ($_.Date) { $_.Date } else { "No Data" } }
                            },
                            @{
                            Label = "CreationTime"
                            Expression = { if ($_.CreationTime) { $_.CreationTime } else { "No Data" } }
                            },
                            @{
                            Label = "LastModifiedTime"
                            Expression = { if ($_.LastModifiedTime) { $_.LastModifiedTime } else { "No Data" } }
                            },
                            @{
                            Label = "Name"
                            Expression = { if ($_.Name) { $_.Name } else { "No Data" } }
                            },
                            @{
                            Label = "FolderPath"
                            Expression = { if ($_.FolderPath) { $_.FolderPath } else { "No Data" } }
                            },
                            @{
                            Label = "FolderId"
                            Expression = { if ($_.FolderId) { $_.FolderId } else { "No Data" } }
                            },
                            @{
                            Label = "Movable"
                            Expression = { if ($_.Movable) { $_.Movable } else { "No Data" } }
                            },
                            @{
                            Label = "VisibleItemsInFolder"
                            Expression = { if ($_.VisibleItemsInFolder) { $_.VisibleItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "HiddenItemsInFolder"
                            Expression = { if ($_.HiddenItemsInFolder) { $_.HiddenItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "ItemsInFolder"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "DeletedItemsInFolder"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "FolderSize"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "FolderAndSubfolderSize"
                            Expression = { if ($_.FolderAndSubfolderSize) { $_.FolderAndSubfolderSize } else { "No Data" } }
                            },
                            @{
                            Label = "OldestItemReceivedDate"
                            Expression = { if ($_.OldestItemReceivedDate) { $_.OldestItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestItemReceivedDate"
                            Expression = { if ($_.NewestItemReceivedDate) { $_.NewestItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestDeletedItemReceivedDate"
                            Expression = { if ($_.OldestDeletedItemReceivedDate) { $_.OldestDeletedItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestDeletedItemReceivedDate"
                            Expression = { if ($_.NewestDeletedItemReceivedDate) { $_.NewestDeletedItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestItemLastModifiedDate"
                            Expression = { if ($_.OldestItemLastModifiedDate) { $_.OldestItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestItemLastModifiedDate"
                            Expression = { if ($_.NewestItemLastModifiedDate) { $_.NewestItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestDeletedItemLastModifiedDate"
                            Expression = { if ($_.OldestDeletedItemLastModifiedDate) { $_.OldestDeletedItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestDeletedItemLastModifiedDate"
                            Expression = { if ($_.NewestDeletedItemLastModifiedDate) { $_.NewestDeletedItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "TopSubjectSize"
                            Expression = { if ($_.TopSubjectSize) { $_.TopSubjectSize } else { "No Data" } }
                            },
                            @{
                            Label = "TopSubjectCount"
                            Expression = { if ($_.TopSubjectCount) { $_.TopSubjectCount } else { "No Data" } }
                            },
                            @{
                            Label = "IsValid"
                            Expression = { if ($_.IsValid) { $_.IsValid } else { "No Data" } }
                            }

            #$UserCalender
            $userMailbox = get-mailboxstatistics -Identity $User.UserPrincipalName | Select DisplayName,DeletedItemCount,ItemCount,TotalDeletedItemSize,TotalItemSize,MessageTableTotalSize,MessageTableAvailableSize,AttachmentTableTotalSize,AttachmentTableAvailableSize,OtherTablesTotalSize,OtherTablesAvailableSize,IsEncrypted,LastInteractionTime,LastLoggedOnUserAccount,LastLogoffTime,LastLogonTime,AssociatedItemCount ,IsValid,IsArchiveMailBox,MailBoxType 
            #$userMailbox
            # Collect Calendar information
            if ($UserCalender.length -eq 1)
            {
                $CalendarResults += [pscustomobject]$UserCalender
            }
            else
            {
                ForEach ($obj in $UserCalender)
                {
                    $CalendarResults += [pscustomobject]$obj
                }
            }
            #write-host $CalendarResults
            # Collect Mailbox information
            if ($userMailbox.length -eq 1)
            {
                $MailboxResults += [pscustomobject]$userMailbox
            }
            else
            {
                ForEach ($mailBoxObj in $userMailbox)
                {
                    $MailboxResults += [pscustomobject]$mailBoxObj
                }
            }
            #write-host $MailboxResults
        }
        $CalendarResults | Export-CSV  $CalendarDataPath
        $MailboxResults | Export-CSV  $MailBoxDataPath
    }
    else
    {
    #No input file found, gather all mailboxes from Office 365
        $objUsers = get-mailbox  -ResultSize Unlimited | get-mailboxstatistics | Select DisplayName,DeletedItemCount,ItemCount,TotalDeletedItemSize,TotalItemSize,MessageTableTotalSize,MessageTableAvailableSize,AttachmentTableTotalSize,AttachmentTableAvailableSize,OtherTablesTotalSize,OtherTablesAvailableSize,IsEncrypted,LastInteractionTime,LastLoggedOnUserAccount,LastLogoffTime,LastLogonTime,AssociatedItemCount ,IsValid,IsArchiveMailBox,MailBoxType | Export-CSV  $MailBoxDataPath
        $objUsersCal = get-mailbox  -ResultSize Unlimited | Get-MailboxFolderStatistics  –FolderScope 'Calendar' | 
        Select-object   @{
                            Label = "Identity"
                            Expression = { if ($_.Identity) { $_.Identity } else { "No Data" } }
                            },
                            @{
                            Label = "Date"
                            Expression = { if ($_.Date) { $_.Date } else { "No Data" } }
                            },
                            @{
                            Label = "CreationTime"
                            Expression = { if ($_.CreationTime) { $_.CreationTime } else { "No Data" } }
                            },
                            @{
                            Label = "LastModifiedTime"
                            Expression = { if ($_.LastModifiedTime) { $_.LastModifiedTime } else { "No Data" } }
                            },
                            @{
                            Label = "Name"
                            Expression = { if ($_.Name) { $_.Name } else { "No Data" } }
                            },
                            @{
                            Label = "FolderPath"
                            Expression = { if ($_.FolderPath) { $_.FolderPath } else { "No Data" } }
                            },
                            @{
                            Label = "FolderId"
                            Expression = { if ($_.FolderId) { $_.FolderId } else { "No Data" } }
                            },
                            @{
                            Label = "Movable"
                            Expression = { if ($_.Movable) { $_.Movable } else { "No Data" } }
                            },
                            @{
                            Label = "VisibleItemsInFolder"
                            Expression = { if ($_.VisibleItemsInFolder) { $_.VisibleItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "HiddenItemsInFolder"
                            Expression = { if ($_.HiddenItemsInFolder) { $_.HiddenItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "ItemsInFolder"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "DeletedItemsInFolder"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "FolderSize"
                            Expression = { if ($_.ItemsInFolder) { $_.ItemsInFolder } else { "No Data" } }
                            },
                            @{
                            Label = "FolderAndSubfolderSize"
                            Expression = { if ($_.FolderAndSubfolderSize) { $_.FolderAndSubfolderSize } else { "No Data" } }
                            },
                            @{
                            Label = "OldestItemReceivedDate"
                            Expression = { if ($_.OldestItemReceivedDate) { $_.OldestItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestItemReceivedDate"
                            Expression = { if ($_.NewestItemReceivedDate) { $_.NewestItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestDeletedItemReceivedDate"
                            Expression = { if ($_.OldestDeletedItemReceivedDate) { $_.OldestDeletedItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestDeletedItemReceivedDate"
                            Expression = { if ($_.NewestDeletedItemReceivedDate) { $_.NewestDeletedItemReceivedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestItemLastModifiedDate"
                            Expression = { if ($_.OldestItemLastModifiedDate) { $_.OldestItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestItemLastModifiedDate"
                            Expression = { if ($_.NewestItemLastModifiedDate) { $_.NewestItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "OldestDeletedItemLastModifiedDate"
                            Expression = { if ($_.OldestDeletedItemLastModifiedDate) { $_.OldestDeletedItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "NewestDeletedItemLastModifiedDate"
                            Expression = { if ($_.NewestDeletedItemLastModifiedDate) { $_.NewestDeletedItemLastModifiedDate } else { "No Data" } }
                            },
                            @{
                            Label = "TopSubjectSize"
                            Expression = { if ($_.TopSubjectSize) { $_.TopSubjectSize } else { "No Data" } }
                            },
                            @{
                            Label = "TopSubjectCount"
                            Expression = { if ($_.TopSubjectCount) { $_.TopSubjectCount } else { "No Data" } }
                            },
                            @{
                            Label = "IsValid"
                            Expression = { if ($_.IsValid) { $_.IsValid } else { "No Data" } }
                            }| Export-CSV  $CalendarDataPath
    }

    write-host "Finished"
    write-host "Calendar file created: $CalendarDataPath"
    write-host "Mailbox file created: $MailBoxDataPath"



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
[Security.SecureString]$Office365AdminPassword
)

#Encrypt password for transmission to Office365
#$SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force
$SecureOffice365Password=$Office365AdminPassword
#Build credentials object
#write-host $Office365AdminPassword
$Office365Credentials = New-Object System.Management.Automation.PSCredential $Office365AdminUsername, $SecureOffice365Password
#Create remote Powershell session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Office365credentials -Authentication Basic –AllowRedirection
#Import the session
Import-PSSession $Session -AllowClobber | Out-Null
}
# Start script
. Main



