home
 Function home{
 cls
    
    $global:choice = Read-Host -Prompt "

    *****************************************************
    |                                                   |
    |______________ ACTIVE DIRECTORY ___________________|
    |                                                   |
    |1. Export Members from SG                          |  
    |2. BGV Check                                       |
    |3. Export Members from Bulk SG                     |                                         
    |                                                   |
    |                                                   |
    |______________ MICROSOFT EXCHANGE _________________|
    |                                                   |
    |4. Export Members from DL                          |
    |5. Export Members from DDL                         |
    |6. Add Bulk users to DL                            |
    |7. Add user to DL                                  |
    |8. Check user is a part of DDL                     |
    |0. Exit                                            |
    |                                                   |
    |                                                   |
    |               SANJAY KRISHNAN JV                  | 
    |                  My Scripts                       |
    *****************************************************
    
    Please Select the Action from the above"
    container
}

Function container{

#####################Export Members from SG################################
    if($global:choice -eq 1){
        cls
        Import-Module ActiveDirectory

        #Array to store data
        $results = @()

        # Define the security group name
        $groupName = Read-Host -Prompt "Please Enter the SG Name"

        cls

        Write-Host "Exporting details of users from the SG:" -f red -nonewline; Write-Host "groupName" -f green;


        # Get the members of the specified security group
        $groupMembers = Get-ADGroupMember -server dc.server.com -Identity $groupName -Recursive | Where-Object { $_.objectClass -eq 'user' }

        # Loop through each group member and get their manager's name
        foreach ($member in $groupMembers) {
            #User Details
            $user = Get-ADUser  -server dc.server.com -Identity $member.SamAccountName -Property DisplayName, Manager, EmailAddress,  accessToSensitiveData, backgroundCheckLastCompleted, department, Enabled
            #Manager Details
            $manager = Get-ADUser -server dc.server.com -Identity $user.Manager -Property DisplayName
    
            #Writing  details to row in the array
            $result =[PSCustomObject]@{
                UserName = $user.DisplayName
                EmailAddress = $user.EmailAddress
                ManagerName = $manager.DisplayName
                AccessToSensitiveData = $user.accessToSensitiveData
                LastBGVDate = $user.backgroundCheckLastCompleted
                Department = $user.department
                Enabled = $user.Enabled
            }
            #adding each result to new raw in the array
            $results += $Result   
        }
        #Export final details to the csv file
        $results | Export-Csv -Path "C:\new\SGexported.csv" -NoTypeInformation

        cls

        Write-Output "Exported $($results.Length) data to myfilw.csv Successfully"

        # Define the path to the CSV file
       $csvFilePath = "C:\new\SGexported.csv"

        # Define the path to the Excel executable (usually not needed if Excel is in the system PATH)
       $excelPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

        # Open the CSV file in Excel
       Start-Process -FilePath $excelPath -ArgumentList $csvFilePath -Verb runasuser

        Read-Host -Prompt "Press Any Key to Continue"
        home
    }

    
    #####################BGV Check################################

    elseif ($global:choice -eq 2){
        Import-Module ActiveDirectory
        $username = Read-Host -Prompt "Please Enter the username" 
        echo " "
        get-aduser $username -server dc.server.com -pr SamAccountName,backgroundCheckLastCompleted,accessToSensitiveData
        
        Read-Host -Prompt "Press any key to continue.."
        home
    }
     
    #####################Export Members from Bulk SG################

    elseif ($global:choice -eq 3){
        Import-Module ActiveDirectory
        # Function to sanitize worksheet names
function Sanitize-WorksheetName {
    param (
        [string]$name
    )
    # Remove invalid characters
    $name = $name -replace '[\\/*?:\[\]]', ''
    # Truncate to 31 characters
    if ($name.Length -gt 31) {
        $name = $name.Substring(0, 31)
    }
    return $name
}

# Path to the text file containing the list of security groups
$groupFilePath = "C:\scripts\sgbulk.txt"

# Read the security groups from the text file
$groups = Get-Content -Path $groupFilePath

# Path to the Excel file to export the data
$excelFilePath = "C:\scripts\ExportedUsers.xlsx"

# Initialize a counter for progress tracking
$totalGroups = $groups.Count
$currentGroup = 0

# Loop through each group and get the users
foreach ($group in $groups) {
    $currentGroup++
    Write-Progress -Activity "Exporting Users" -Status "Processing group $currentGroup of $totalGroups" -PercentComplete (($currentGroup / $totalGroups) * 100)
    
    $groupMembers = Get-ADGroupMember  -server dc.server.com -Identity $group -Recursive | Where-Object { $_.objectClass -eq 'user' }
    $userInfo = @()
    foreach ($member in $groupMembers) {
        $user = Get-ADUser -server dc.server.com -Identity $member.SamAccountName -Properties DisplayName, Enabled
        $userInfo += [PSCustomObject]@{
            UserName    = $user.SamAccountName
            DisplayName = $user.DisplayName
            Enabled     = $user.Enabled
        }
    }
    
    # Sanitize the worksheet name
    $worksheetName = Sanitize-WorksheetName -name $group
    
    # Export the user information to a worksheet in the Excel file
    $userInfo | Export-Excel -Path $excelFilePath -WorksheetName $worksheetName
}

Write-Output "User information exported successfully to $excelFilePath"


    }
#####################Export Members from DL################################
    elseif($global:choice -eq 4){
        cls

        $DistributionList = Read-Host -Prompt "Please Enter the Email address of the DL" # Group-Name or Group Email
        cls
        Write-Host "Exporting Members of the DL $DistributionList"
         cls
        $CSVFilePath = "C:\scripts\$DistributionList.csv"
        Try {
           #Connect to Exchange Online
           Connect-ExchangeOnline -ShowBanner:$False
           #Get Distribution List Members and Exports to CSV
           Get-DistributionGroupMember -Identity $DistributionList -ResultSize Unlimited | Select DisplayName, PrimarySMTPAddress, Alias, Department, Manager | Export-Csv $CSVFilePath -NoTypeInformation
        }
        Catch {
           write-host -f Red "Error:" $_.Exception.Message
        }
        cls
        Write-Host "Exported the list of members of DL $DistributionList to the location $CSVFilePath"
        cls
       
        }


#####################Export Members from DDL################################
    elseif ($global:choice -eq 5){

        cls
        $DistributionList = Read-Host -Prompt "Please Enter the name of the DDL" # Group-Name or Group Email
        cls
        Write-Host "Exporting Members of the DL $DistributionList"
        cls
        $CSVFilePath = "C:\scripts\$DistributionList.csv"
        Try {
            #Connect to Exchange Online
            Connect-ExchangeOnline -ShowBanner:$False
            #Get Distribution List Members and Exports to CSV
            $DDL= Get-DynamicDistributionGroup "$DistributionList"
        Get-Recipient -RecipientPreviewFilter $DDL.RecipientFilter -OrganizationalUnit $DDL.RecipientContainer -ResultSize Unlimited | Export-CSV -Path c:\scripts\DDL1.csv
        }
        Catch {
            write-host -f Red "Error:" $_.Exception.Message
        }
        cls
        Write-Host "Exported the list of members of DL $DistributionList"
        
    }

    
#####################Add Bulk users to DL################################
    elseif ($global:choice -eq 6){

        cls
        $GroupEmailID = Read-Host -Prompt "Please Enter the Email of the DL Name" # Group Email
        cls
        $CSVFile  = Read-Host -Prompt "Please Enter the Path of the CSV File" #Copy the path as CopyAsPath
        
        #Connect to Exchange Online
        Connect-ExchangeOnline -ShowBanner:$False
 
        #Get Existing Members of the Distribution List
        $DLMembers =  Get-DistributionGroupMember -Identity $GroupEmailID -ResultSize Unlimited | Select -Expand PrimarySmtpAddress
 
        #Import Distribution List Members from CSV
        Import-CSV $CSVFile -Header "UPN" | ForEach {
            #Check if the Distribution List contains the particular user
            If ($DLMembers -contains $_.UPN)
            {
                Write-host -f Yellow "User is already member of the Distribution List:"$_.UPN
            }
            Else
            {       
                Add-DistributionGroupMember -Identity $GroupEmailID -Member $_.UPN
                Write-host -f Green "Added User to Distribution List:"$_.UPN
            }
        }
    }


    #####################Add Bulk users to DL################################

    elseif ($global:choice -eq 7){

        Connect-ExchangeOnline -ShowBanner:$False 
        $DLName = Read-Host -Prompt "Please Enter the DL Name" # Group Email
        $Count = Read-Host -Prompt "Please Enther the Number of users need to be added" # Group Email
  
        for($i= 0; $i -lt $Count; $i++){

            $UserEmail = Read-Host -Prompt "Please Enter Email to add as a member of $DLName" # Group Email

            # Check if the user is already a member of the DL
            $member = Get-DistributionGroupMember -Identity $DLName | Where-Object { $_.PrimarySmtpAddress -eq $UserEmail }

            if ($member) {
                 Write-host  -f Red "User $UserEmail is already a member of $DLName."
            } else {
                # Add the user to the DL
                Add-DistributionGroupMember -Identity $DLName -Member $UserEmail
                 Write-host  -f Green "User $UserEmail has been added to $DLName."
            }
        }
        Read-Host -Prompt "Press any key to continue.."
    }

        #####################Check user is a part of DDL################################

    elseif ($global:choice -eq 8){

        Connect-ExchangeOnline -ShowBanner:$False
        $userEmail = Read-Host -Prompt "Please Enter the user Email addresss" # User Email
        $DDLName = Read-Host -Prompt "Please Enther DDL Name" # DDL Name
  
        $DDL = Get-DynamicDistributionGroup -Identity $DDLName

        # Get the filter applied to the Dynamic Distribution Group
        $filter = $DDL.RecipientFilter

        # Check if the user matches the filter
        $user = Get-Recipient -Filter $filter | Where-Object { $_.PrimarySmtpAddress -eq $userEmail }

        if ($user) {
            Write-host  -f Green "
    
            $userEmail is a member of $DDLName."
        } else {
            Write-host  -f Red "
    
            $userEmail is not a member of $DDLName."
        }

        Read-Host -Prompt "Press any key to continue.."
    }

     elseif ($global:choice -eq 0){

       exit
    }

    else{
    write-Host "Select valid choice"
    
    }
    home
}
 
 