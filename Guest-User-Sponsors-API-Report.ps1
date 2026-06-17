# ================================
# Connect to Graph (Managed Identity)
# ================================

Connect-MgGraph -Identity

# ================================
# Config
# ================================

$senderEmail = "dean.hoile@deansbean2.com"
$NotifyDays  = 30
$Today       = (Get-Date).Date

Write-Output "Retrieving guest users..."

# ================================
# Get Guests
# ================================

[array]$Guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property Id,DisplayName,Mail,EmployeeHireDate

# ================================
# Cache for sponsor lookups (performance)
# ================================

$SponsorCache = @{}

# ================================
# Group guests per sponsor
# ================================

$SponsorGuestMap = @{}

foreach ($Guest in $Guests) {

    if (!$Guest.EmployeeHireDate) { continue }

    $HireDate      = (Get-Date $Guest.EmployeeHireDate).Date
    $DaysUntilHire = ($HireDate - $Today).Days

    if ($DaysUntilHire -lt 0 -or $DaysUntilHire -gt $NotifyDays) { continue }

    # ================================
    # RAW API: Get sponsors
    # ================================

    $uri = "https://graph.microsoft.com/v1.0/users/$($Guest.Id)/sponsors"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    }
    catch {
        Write-Output "Failed to retrieve sponsors for $($Guest.DisplayName)"
        continue
    }

    if (!$response.value) { continue }

    foreach ($s in $response.value) {

        $SponsorId = $s.id

        # ================================
        # Cache lookup (avoid duplicate calls)
        # ================================

        if ($SponsorCache.ContainsKey($SponsorId)) {
            $SponsorUser = $SponsorCache[$SponsorId]
        }
        else {
            try {
                $SponsorUser = Get-MgUser `
                    -UserId $SponsorId `
                    -Property DisplayName,Mail,UserPrincipalName

                $SponsorCache[$SponsorId] = $SponsorUser
            }
            catch {
                Write-Output "Failed to resolve sponsor ID $SponsorId"
                continue
            }
        }

        # ================================
        # Extract sponsor details
        # ================================

        $SponsorEmail = $SponsorUser.Mail
        if (!$SponsorEmail) {
            $SponsorEmail = $SponsorUser.UserPrincipalName
        }

        $SponsorName = $SponsorUser.DisplayName

        # Validate email
        if (!$SponsorEmail -or $SponsorEmail -notmatch "@") { continue }

        # ================================
        # Build mapping
        # ================================

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

# ================================
# No results check
# ================================

if ($SponsorGuestMap.Count -eq 0) {
    Write-Output "No expiring guests found"
    Write-Output "Process complete."
    return
}

Write-Output "Preparing emails..."

# ================================
# Send emails per sponsor
# ================================

foreach ($SponsorEmail in $SponsorGuestMap.Keys) {

    $Sponsor   = $SponsorGuestMap[$SponsorEmail]
    $GuestList = $Sponsor.Guests | Sort-Object ReviewDate

    # ================================
    # Build HTML rows
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

    Write-Output "Email sent to $SponsorEmail for $($GuestList.Count) guests"
}

Write-Output "Process complete."
