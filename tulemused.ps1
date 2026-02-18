# 1. Define the Style (CSS) for the HTML page
$Header = @"
<style>
    body { font-family: Calibri, sans-serif; }
    table { border-collapse: collapse; width: 600px; }
    th { background-color: #4CAF50; color: white; padding: 10px; text-align: left; }
    td { border: 1px solid #ddd; padding: 8px; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .DONE { color: green; font-weight: bold; }
    .MISSING { color: red; font-weight: bold; }
</style>
"@

$tulemused = New-Object System.Collections.Generic.List[PSCustomObject]

function Lisa-Kontroll {
    param($Nimi, $Tingimus)
    # We use English status here for the HTML logic
    $Staatus = if ($Tingimus) { "DONE" } else { "MISSING" }
    $tulemused.Add([PSCustomObject]@{ Checkpoint = $Nimi; Status = $Staatus })
}

$IPCheck = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -match "^10\.0\..*\.10$" }).Count -gt 0

# --- START CHECKS ---
Lisa-Kontroll "Server Name is AD1" ($env:COMPUTERNAME -eq "AD1")
Lisa-Kontroll "IP is 10.0.x.10" ($IPCheck)
Lisa-Kontroll "Domain perenimi.local" (Test-Connection -ComputerName "KRUTTO.LOCAL" -Count 1 -Quiet -ErrorAction SilentlyContinue)

# OUs
foreach ($ou in "KASUTAJAD", "LEKTORID", "TUDENGID") {
    Lisa-Kontroll "OU $ou exists" (Get-ADOrganizationalUnit -Filter "Name -eq '$ou'" -ErrorAction SilentlyContinue)
}

# Roles
foreach ($roll in "AD-Domain-Services", "DHCP", "DNS") {
    Lisa-Kontroll "Role $roll installed" (Get-WindowsFeature -Name $roll).Installed
}

# Bonus: License
$slmgr = Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" | Select-Object -First 1
$Days = [math]::Round($slmgr.RemainingGracePeriod / 1440, 1)
Lisa-Kontroll "License Days Left: $Days" ($Days -gt 0)
# --- END CHECKS ---

# 2. Generate the HTML File
$FilePath = "$env:USERPROFILE\Desktop\tulemused.html"

$tulemused | ConvertTo-Html -Head $Header -Title "Course Progress Report" | Out-File $FilePath

# 3. Open the file automatically
Invoke-Item $FilePath

Write-Host "Report generated on your Desktop: Lab_Report.html" -ForegroundColor Cyan