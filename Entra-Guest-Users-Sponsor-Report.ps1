# Log into Entra ID with managed identity
Connect-MgGraph -Identity

# Sender account
$senderEmail = "corporate.alerts@boggle.com"

# Config
$NotifyDays = 30
$Today = (Get-Date).Date

Write-Output "Retrieving guest users..." -ForegroundColor Green

[array]$Guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property Id,DisplayName,Mail,EmployeeHireDate,Sponsors `
    -ExpandProperty Sponsors

Write-Output "Total guests retrieved: $($Guests.Count)"

# Show a sample of guests
$Guests | Select-Object -First 5 DisplayName, Mail, EmployeeHireDate | Format-Table


# ================================
# Group guests per sponsor
# ================================

$SponsorGuestMap = @{}

foreach ($Guest in $Guests) {

    if (!$Guest.EmployeeHireDate) { continue }

    $HireDate = (Get-Date $Guest.EmployeeHireDate).Date
    $DaysUntilHire = ($HireDate - $Today).Days

    if ($DaysUntilHire -ge 0 -and $DaysUntilHire -le $NotifyDays) {

        if (!$Guest.Sponsors) { continue }

        foreach ($Sponsor in $Guest.Sponsors) {

            $SponsorEmail = $Sponsor.additionalProperties.mail
            $SponsorName  = $Sponsor.additionalProperties.displayName

            if (!$SponsorEmail) { continue }
            if ($SponsorEmail -notmatch "@") { continue }

            if (-not $SponsorGuestMap.ContainsKey($SponsorEmail)) {
                $SponsorGuestMap[$SponsorEmail] = @{
                    SponsorName = $SponsorName
                    Guests      = @()
                }
            }

            $SponsorGuestMap[$SponsorEmail].Guests += [PSCustomObject]@{
                Name       = $Guest.DisplayName
                Email      = $Guest.Mail
                ReviewDate = $HireDate
                DaysLeft   = $DaysUntilHire
            }
        }
    }
}

# ================================
# NO RESULTS CHECK (IMPORTANT)
# ================================

if ($SponsorGuestMap.Count -eq 0) {
    Write-Output "No expiring guests found" -ForegroundColor Yellow
    Write-Output "Process complete." -ForegroundColor Green
    return
}

Write-Output "Preparing emails..." -ForegroundColor Green

# ================================
# Send emails per sponsor
# ================================

foreach ($SponsorEmail in $SponsorGuestMap.Keys) {

    $Sponsor   = $SponsorGuestMap[$SponsorEmail]
    $GuestList = $Sponsor.Guests | Sort-Object ReviewDate


    # ================================
    # Build rows
    # ================================

    $rows = @()

    foreach ($Guest in $GuestList) {

        $rows += @"
<tr>
    <td>$($Guest.Name)</td>
    <td>$($Guest.Email)</td>
    <td>$($Guest.ReviewDate.ToString("dd/MM/yyyy"))</td>
    <td style="text-align:center;">$($Guest.DaysLeft)</td>
</tr>
"@
    }

    $rowsHtml = $rows -join "`n"

    
    # ================================
    # HTML Email Body
    # ================================


    $html = @"
<html>
<head>
<style>
body { font-family: Arial, sans-serif; font-size: 14px; color: #333; }
h2 { color: #0078D4; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; }
th { background-color: #0078D4; color: white; padding: 10px; text-align: left; }
td { padding: 8px; border-bottom: 1px solid #ddd; }
tr:nth-child(even) { background-color: #f9f9f9; }
.footer { margin-top: 20px; font-size: 12px; color: #666; }
</style>
</head>
<body>

<h2>Guest Access Review Required</h2>

<p>Dear $($Sponsor.SponsorName),</p>

<p>The following guest accounts you sponsor are due for review within the next <b>$NotifyDays days</b>.</p>

<table>
<tr>
    <th>Name</th>
    <th>Email</th>
    <th>Review Date</th>
    <th>Days Left</th>
</tr>

$rowsHtml

</table>

<p>Please review whether these users still require access, if still required please raise an OAA ticket to extend their access.</p>

<p><b>This message was sent from an unmonitored email address.</b></p>

<div class="footer">
Generated on $(Get-Date -Format "dd/MM/yyyy")
</div>

</body>
</html>
"@

    
    # ================================
    # Send Email
    # ================================

    Send-MgUserMail -UserId $senderEmail -Message @{
        Subject = "Guest Access Review Required"
        Body = @{
            ContentType = "HTML"
            Content     = $html
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = $SponsorEmail
                }
            }
        )
    }

    Write-Output "Email sent to $SponsorEmail for $($GuestList.Count) guests" -ForegroundColor Cyan
}

Write-Output "Process complete." -ForegroundColor Green