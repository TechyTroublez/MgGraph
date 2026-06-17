# ================================
# DRY RUN CONFIG
# ================================
$OutputFile = "/Users/dean.hoile/GuestDryRun.html"
$NotifyDays = 365
$Today = (Get-Date).Date

Write-Host "Retrieving guest users..." -ForegroundColor Green

[array]$Guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property Id,DisplayName,Mail,EmployeeHireDate

Write-Host "Total guests retrieved: $($Guests.Count)" -ForegroundColor Yellow

# Group guests per sponsor
$SponsorGuestMap = @{}

foreach ($Guest in $Guests) {

    if (!$Guest.EmployeeHireDate) { continue }

    $HireDate = (Get-Date $Guest.EmployeeHireDate).Date
    $DaysUntilHire = ($HireDate - $Today).Days

    if ($DaysUntilHire -ge 0 -and $DaysUntilHire -le $NotifyDays) {

        # ================================
        # Get sponsors reliably
        # ================================
        $Sponsors = Get-MgUserSponsor -UserId $Guest.Id

        if (!$Sponsors) {
            Write-Host "No sponsor for $($Guest.DisplayName)" -ForegroundColor DarkYellow
            continue
        }

        foreach ($Sponsor in $Sponsors) {

            # Get sponsor details
            $SponsorDetails = Get-MgUser -UserId $Sponsor.Id -Property DisplayName,Mail

            $SponsorEmail = $SponsorDetails.Mail
            $SponsorName  = $SponsorDetails.DisplayName

            if (!$SponsorEmail) {
                Write-Host "Sponsor has no email: $SponsorName" -ForegroundColor DarkYellow
                continue
            }

            if (!$SponsorGuestMap.ContainsKey($SponsorEmail)) {
                $SponsorGuestMap[$SponsorEmail] = @{
                    SponsorName = $SponsorName
                    Guests = @()
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

Write-Host "Sponsors mapped: $($SponsorGuestMap.Keys.Count)" -ForegroundColor Yellow

Write-Host "Building HTML preview..." -ForegroundColor Yellow

# ================================
# Build HTML rows
# ================================
$rows = foreach ($SponsorEmail in $SponsorGuestMap.Keys) {

    $Sponsor = $SponsorGuestMap[$SponsorEmail]

    foreach ($Guest in $Sponsor.Guests) {

@"
<tr>
    <td>$($Sponsor.SponsorName)</td>
    <td>$($Guest.Name)</td>
    <td>$($Guest.Email)</td>
    <td>$($Guest.ReviewDate.ToString("dd/MM/yyyy"))</td>
    <td>$($Guest.DaysLeft)</td>
</tr>
"@
    }
}

# Count total guests
$totalGuests = ($SponsorGuestMap.Values | ForEach-Object { $_.Guests.Count } | Measure-Object -Sum).Sum

# ================================
# HTML Page
# ================================
$html = @"
<html>
<head>
<style>
body { font-family: Arial; }
table { border-collapse: collapse; width: 100%; }
th { background-color: #0078D4; color: white; padding: 8px; }
td { padding: 8px; border-bottom: 1px solid #ddd; }
tr:nth-child(even) { background-color: #f2f2f2; }
</style>
</head>
<body>

<h2>DRY RUN - Guest Access Review Report</h2>

<p><b>Total Guests:</b> $totalGuests</p>
<p><b>Note:</b> This is a preview only. No emails have been sent.</p>

<table>
<tr>
    <th>Sponsor</th>
    <th>Guest Name</th>
    <th>Guest Email</th>
    <th>Review Date</th>
    <th>Days Left</th>
</tr>

$rows

</table>

<br>

<p>Generated: $(Get-Date -Format "dd/MM/yyyy")</p>

</body>
</html>
"@

# ================================
# Save + Open
# ================================

$html | Out-File -FilePath $OutputFile -Encoding utf8

Write-Host "Dry run report saved to $OutputFile" -ForegroundColor Cyan