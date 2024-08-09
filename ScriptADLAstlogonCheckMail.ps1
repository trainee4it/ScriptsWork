



function Get-LatestLogon
{
      <#
    .SYNOPSIS
    Retrieves the latest logon information for a specified user from all domain controllers in a specified site.

    .DESCRIPTION
    The `Get-LatestLogon` function searches for the latest logon time of a specified user across all domain controllers in a specified site. It returns the logon information if the latest logon is within the specified number of days.

    .PARAMETER UserName
    Specifies the username to search for.
    - Type: String
    - Required: True
    - Accept pipeline input: True (ByValue)
    - Accept wildcard characters: False

    .PARAMETER DaysAgo
    Specifies the number of days ago to use as a reference date. The function will check if the latest logon is within this period.
    - Type: Int
    - Required: True
    - Accept pipeline input: False
    - Accept wildcard characters: False

    .PARAMETER SiteName
    Specifies the site name to filter domain controllers.
    - Type: String
    - Required: True
    - Accept pipeline input: False
    - Accept wildcard characters: False

    .OUTPUTS
    System.Management.Automation.PSObject
    The function returns a custom object with the following properties:
    - dc: The domain controller where the latest logon was recorded.
    - date: The date and time of the latest logon.
    - samaccountname: The SAM account name of the user.
    - enabled: Indicates whether the user account is enabled.
    - UserPrincipalname: The user principal name of the user.

    .EXAMPLE
    Get-LatestLogon -UserName "jdoe" -DaysAgo 30 -SiteName "NYC"
    This command retrieves the latest logon information for the user "jdoe" within the last 30 days from all domain controllers in the "NYC" site.

    .NOTES
    Ensure you have the necessary permissions to query Active Directory and access domain controllers.
    The function converts the last logon timestamp to a readable date format.

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-aduser
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-addomaincontroller
    #>

    [CmdletBinding()]
    param(
    
    [Parameter(Mandatory= $true , ValueFromPipeline = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [int]$DaysAgo,
    [Parameter(Mandatory = $true)]
    [string]$siteName
    
    )
    # Specify the username to search for
    #$usernames = "carmen-hoofs"

    $ReferenceDate = (Get-Date).AddDays(-$DaysAgo)

    # Get the list of all Domain Controllers
   
    
    $DCs = Get-ADDomainController -Filter * | where {$_.site -eq $SiteName}
    
    # Initialize a variable to hold the latest logon time
    $latestLogon = $null

    # Loop through all domain controllers
    foreach($DC in $DCs)
    {
    # Get the last logon time stamp
      
        $user = Get-ADUser $username -Server $DC.HostName -Properties lastLogon

        # If the user's last logon time is later than the current latest logon time, update the latest logon time
        if($user.lastLogon -gt $latestLogon)
        {
            $latestLogon = $user.lastLogon
            $LatestDc = $dc.HostName
            $SamaccountName = $user.SamAccountName
            $UserPrincipalName = $user.UserPrincipalName
            $Enabled = $user.enabled

        }

        
    }
    # Convert to readable format :)
    $latestLogon = [DateTime]::FromFileTime($latestlogon)
    
    if($latestLogon -lt $ReferenceDate)
    {
        
        #create psobject and output it, please add if something is missing
         
        $props = @{'dc'=$LatestDc;'date'=$latestLogon;'samaccountname'=$SamaccountName;'enabled'=$Enabled;'UserPrincipalname'=$UserPrincipalName}
        $object = New-Object -TypeName psobject -Property $props
        $object

    }
    else
    {

        $object


    }
  
    

}


function Get-MailboxStatus
{
<#
.SYNOPSIS
Retrieves the mailbox status of a specified user.

.DESCRIPTION
The `Get-MailboxStatus` function checks if a user has a mailbox in Exchange Online by querying their User Principal Name (UPN). If the user has a mailbox, the function returns the mailbox information. If the user does not have a mailbox, the function returns null.

.PARAMETER UserPrincipalName
Specifies the User Principal Name (UPN) of the user to check for a mailbox.
- Type: String
- Required: True
- Accept pipeline input: True (ByValue)
- Accept wildcard characters: False

.OUTPUTS
System.Management.Automation.PSObject
The function returns the mailbox information if the user has a mailbox, or null if the user does not have a mailbox.

.EXAMPLE
Get-MailboxStatus -UserPrincipalName "user@example.com"
This command retrieves the mailbox status for the user with the UPN "user@example.com".

.NOTES
Ensure you have the necessary permissions to query Exchange Online.
#>


[CmdletBinding()]
    param(
    
    [Parameter(Mandatory= $true , ValueFromPipeline = $true)]
    [string]$UserPrincipalName
    )
  
#Does user have a mailbox ?
        
        
        try
        {
            
            $MailboxUser  = Get-EXOMailbox -UserPrincipalName $UserPrincipalName -ErrorAction Stop
            

        }
        catch
        {


            $MailboxUser = $null
            

        }



        Write-Output -InputObject $MailboxUser






}


#MAIN we start here
<#
.SYNOPSIS
Generates a report of user logon information and mailbox permissions.

.DESCRIPTION
This script connects to Exchange Online and Active Directory to gather information about user logons and mailbox permissions. It performs the following tasks:
1. Connects to Exchange Online if not already connected.
2. Retrieves all mailboxes and their permissions from Exchange Online.
3. Retrieves all users from Active Directory who match a specific distinguished name pattern.
4. For each user, it checks the latest logon information and mailbox status.
5. Compiles a report with the user's logon date, mailbox status, and any additional mailbox permissions.
6. Outputs the report in a readable format.

.PARAMETER DaysAgo
Specifies the number of days ago to use as a reference date for logon information.

.NOTES
Ensure you have the necessary permissions to query Exchange Online and Active Directory.
#>

# Initialize an array to hold the report data


$ReportArray = New-Object -TypeName System.Collections.ArrayList

$DaysAgo = 90
if(!(Get-ConnectionInformation))
{

    Connect-ExchangeOnline
    

}
if (!($MailBoxes)) {
    $MailBoxes = @()
    $allMailboxes = Get-EXOMailbox -ResultSize Unlimited
    $totalMailboxes = $allMailboxes.Count
    $currentMailbox = 0

    foreach ($mailbox in $allMailboxes) {
        $currentMailbox++
        Write-Progress -Activity "Retrieving Mailbox Permissions" -Status "Processing mailbox $currentMailbox of $totalMailboxes" -PercentComplete (($currentMailbox / $totalMailboxes) * 100)

        $mailboxUserPrincipalName = $mailbox.UserPrincipalName
        $mailboxType = $mailbox.RecipientTypeDetails
        $permissions = Get-EXOMailboxPermission -Identity $mailboxUserPrincipalName | Where-Object { $_.User -ne "NT AUTHORITY\SELF" } | Select-Object @{Name="Mailbox";Expression={$mailboxUserPrincipalName}},@{Name="MailboxType";Expression={$mailboxType}}, User, AccessRights, IsInherited, Deny

        $MailBoxes += $permissions
    }
}


if(!$users)
{
    
    $Users = Get-ADUser -Properties lastlogon -Filter * | Where-Object {$psitem.distinguishedname -like '*uitdienst*'}
}   

$CountUsersToProcess = $users.count

# Process each user to gather logon and mailbox information

foreach($User in $Users)
{
    $USerObject = $null
    $UserObject =Get-LatestLogon -User $user.SamAccountName -DaysAgo $DaysAgo -siteName 'default-first-site-name'
   # Write-Verbose -Message "Processing... $count users to go" -Verbose
    Write-Host "Processing... $CountUsersToProcess users to go" -ForegroundColor Yellow -BackgroundColor Black
    #$object
   # Get-Mailbox -Identity $user
    
    
    
    
    if($USerObject)
    {
      
        
        

        if($UserMailbox = Get-MailboxStatus -UserPrincipalName $USerObject.UserPrincipalName)
        {

                $UserPrincipalName = $UserMailbox.UserPrincipalName
            
                $MailboxType = $UserMailbox.Recipienttype
        }
        else
        {

                $UserPrincipalName = 'NODIRECTMAILBOX'
            
                $MailboxType = 'NODIRECTMAILBOX'

        }

        
        $MergeHashTable = [ordered]@{'UserPrincipalname'=$UserObject.UserPrincipalName;
                                  'Date'=$USerObject.Date;
                                  'UserMailBox'=$UserPrincipalName;
                                  'MailboxType'=$MailboxType;}
        
        $MailboxOnCall = $null
        
        if($MailboxOnCall = $MailBoxes | Where-Object {$_.user -eq  $USerObject.userprincipalName})
        {
            
            

        
            for ($i = 0; $i -lt $MailboxOnCall.Count; $i++) {
                $MergeHashTable["OtherMailBox$($i + 1)"] = $MailboxOnCall[$i].mailbox
            }
        
            
            

        }
        else
        {


            $MergeHashTable.Add('OtherMailboxesShared', 'NONEDETECTED')
            

        }
        

    
    
        $MergeObject = New-Object -TypeName psobject -Property $MergeHashTable

        $ReportArray.Add($MergeObject) | Out-Null

        #Write-Verbose -Message "$($MergeObject.UserPrincipalName) logged on last $($MergeObject.Date) has its own mailbox $($MergeObject.UserMailBox) and other mailboxes $($MergeObject.OtherMailboxes -join ', ')" -Verbose
    
        $MergeObject | fl
    }

    else
    {



        Write-Verbose -Message "$($User.UserPrincipalName) had a logon less than $($DaysAgo ) " -Verbose


    }

    $CountUsersToProcess--
 
    

 }
 

