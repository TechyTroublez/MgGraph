#Log into Entra ID with managed identity
#Connect-MgGraph -Identity

#This is if you wish to query a number of groups that have set prefix in the name i.e. CORP-01, CORP-02 etc

# Define variables
$groupPrefix = 'corp01' #Add a prefix 
$senderEmail = "send.address@domain.com"
$recipientEmail = "receiver.address@domain.com"
$subject = "Hello"
$body = "Report"

# Initialize an empty array to store CSV content for all groups
$allCsvContent = @()

# Retrieve groups starting with the specified prefix
$groups = Get-MgGroup -Filter "startswith(DisplayName, '$groupPrefix')"

foreach ($group in $groups) {
    $groupId = $group.Id
    $members = Get-MgGroupMember -GroupId $groupId
    $users = @()

    # Retrieve user information and build the report for each group
    foreach ($member in $members) {
        $user = Get-MgUser -UserId $member.Id
        $users += [PSCustomObject]@{ 
            Group = $group.DisplayName
            Name = $user.DisplayName
            USERPRINCIPALNAME = $user.Mail
        }
    }

    # Sort the results by identity
    $sortedUsers = $users | Sort-Object Name

    # Convert to CSV format for this group
    $csvContent = $sortedUsers | ConvertTo-Csv -NoTypeInformation | Out-String

    # Add the CSV content for this group to the array
    $allCsvContent += $csvContent
}

# Combine all CSV content into a single string
$combinedCsvContent = $allCsvContent -join "`r`n"

# Prepare email with all CSVs attached
$params = @{
    message = @{
        subject = $subject
        body = @{
            ContentType = "HTML"
            Content = $body
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = $recipientEmail
                }
            }
        )
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = "group_members_report_all.csv"
                contentBytes = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($combinedCsvContent))
            }
        )
    }
}

# Send email
Send-MgUserMail -UserId $senderEmail -BodyParameter $params