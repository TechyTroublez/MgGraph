# ================================
# Connect to Graph (Managed Identity)
# ================================
#Connect-MgGraph #-Identity

# ================================
# Config
# ================================
$senderEmail   = "dean.hoile@bv-group.com"
$testRecipient = "dean.hoile@bv-group.com"   # 👈 CHANGE THIS
$NotifyDays    = 30
$Today         = (Get-Date).Date

Write-Output "Retrieving guest users..."

# ================================
# Get Guests
# ================================

[array]$Guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property Id,DisplayName,Mail,EmployeeHireDate

# ================================
# Cache for sponsor lookups
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

    # RAW API: Get sponsors
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

        # Cache lookup
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

        $SponsorEmail = $SponsorUser.Mail
        if (!$SponsorEmail) {
            $SponsorEmail = $SponsorUser.UserPrincipalName
        }

        $SponsorName = $SponsorUser.DisplayName

        if (!$SponsorEmail -or $SponsorEmail -notmatch "@") { continue }

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

Write-Output "Preparing TEST email..."

# ================================
# Build ALL rows into ONE email
# ================================

$rows = @()

foreach ($SponsorEmail in $SponsorGuestMap.Keys) {

    $Sponsor = $SponsorGuestMap[$SponsorEmail]

    foreach ($Guest in $Sponsor.Guests) {
        $rows += @"
<tr>
    <td>$($Sponsor.SponsorName)</td>
    <td>$($Guest.Name)</td>
    <td>$($Guest.Email)</td>
    <td>$($Guest.ReviewDate.ToString("dd/MM/yyyy"))</td>
    <td style="text-align:center;">$($Guest.DaysLeft)</td>
</tr>
"@
    }
}

$rowsHtml = $rows -join "`n"

# ================================
# HTML Body
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
</style>
</head>
<body>

<h2>Guest Access Review Report</h2>

<p>The following guest accounts you sponsor are due for review.</p>

<table>
<tr>
    <th>Sponsor</th>
    <th>Guest Name</th>
    <th>Guest Email</th>
    <th>Review Date</th>
    <th>Days Left</th>
</tr>

$rowsHtml

</table>


<p>Please review whether these users still require access, if still required please raise an OAA ticket to extend their access.</p>


<p><b>This message was sent from an unmonitored email address.</b></p>


<div class="footer">
<p>Generated on $(Get-Date -Format "dd/MM/yyyy")</p>

</body>
</html>
"@

# ================================
# Send TEST email
# ================================

Send-MgUserMail -UserId $senderEmail -Message @{
    Subject = "Guest Access Review"
    Body = @{
        ContentType = "HTML"
        Content     = $html
    }
    ToRecipients = @(
        @{
            EmailAddress = @{
                Address = $testRecipient
            }
        }
    )
}

Write-Output "Test email sent to $testRecipient"
Write-Output "Process complete."