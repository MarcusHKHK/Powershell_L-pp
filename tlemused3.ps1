# 1. Funktsioon, mis väljastab staatuse (ühildub kõigi PowerShelli versioonidega)
function Get-Status {
    param ([boolean]$Condition)
    if ($Condition) { return "TEHTUD" } else { return "TEGEMATA" }
}

$RaportObjektid = New-Object System.Collections.Generic.List[PSObject]
$MyDomain = "KRUTTO.LOCAL"

# --- KONTROLLID ---

# Võrk ja Domeen
$Hostname = hostname
$IP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" } | Select-Object -First 1).IPAddress
$DomainExists = Get-ADDomain $MyDomain -ErrorAction SilentlyContinue

$RaportObjektid.Add([PSCustomObject]@{Kontroll = "Serveri nimi on AD1"; Staatus = Get-Status ($Hostname -eq "AD1")})
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "IP on 10.0.XXX.10"; Staatus = Get-Status ($IP -like "10.0.*.10")})
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "Domeen $MyDomain kättesaadav"; Staatus = Get-Status ($null -ne $DomainExists)})

# AD Struktuur
$OUs = @("KASUTAJAD", "LEKTORID", "TUDENGID", "ARVUTID")
foreach ($ou in $OUs) {
    $Exists = Get-ADOrganizationalUnit -Filter "Name -eq '$ou'" -ErrorAction SilentlyContinue
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "OU $ou"; Staatus = Get-Status ($null -ne $Exists)})
}

$Groups = @("Lektorid", "Tudengid", "Salvestamine")
foreach ($grp in $Groups) {
    $Exists = Get-ADGroup -Filter "Name -eq '$grp'" -ErrorAction SilentlyContinue
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "Grupp $grp"; Staatus = Get-Status ($null -ne $Exists)})
}

# Kasutajad
$Users = @("oppejoud1", "oppejoud2", "tudeng1", "tudeng2")
foreach ($u in $Users) {
    $UserObj = Get-ADUser -Identity $u -ErrorAction SilentlyContinue
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "Kasutaja $u"; Staatus = Get-Status ($null -ne $UserObj)})
}

# Rollid
$Roles = @("AD-Domain-Services", "DHCP", "DNS", "WDS", "Web-Server")
foreach ($role in $Roles) {
    $IsInstalled = (Get-WindowsFeature -Name $role -ErrorAction SilentlyContinue).Installed
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "Roll $role"; Staatus = Get-Status ($IsInstalled)})
}

# DHCP Skoop
$DHCPScope = Get-DhcpServerv4Scope -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "HKHK" }
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "DHCP skoop 'HKHK'"; Staatus = Get-Status ($null -ne $DHCPScope)})

# DFS
$Paths = @("F:\DFS_Lektorid", "F:\DFS_Tudengitele")
foreach ($p in $Paths) {
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "Kaust $p"; Staatus = Get-Status (Test-Path $p)})
}
$DFS = Get-DfsnRoot -Path "\\$MyDomain\Tudengid" -ErrorAction SilentlyContinue
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "DFS nimeruum \Tudengid"; Staatus = Get-Status ($null -ne $DFS)})

# GPO ja LAPS
$GPOs = @("7zip", "Chrome", "taustapilt")
foreach ($g in $GPOs) {
    $GPOExists = Get-GPO -Name $g -ErrorAction SilentlyContinue
    $RaportObjektid.Add([PSCustomObject]@{Kontroll = "GPO $g"; Staatus = Get-Status ($null -ne $GPOExists)})
}
$LAPS = Get-Command -Module Microsoft.Windows.LAPS -ErrorAction SilentlyContinue
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "LAPS moodul"; Staatus = Get-Status ($null -ne $LAPS)})

# Boonus: Litsents
$License = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq "55c282d8-bf17-464f-9561-fa08d992f17d" }
$DaysLeft = if ($License) { [Math]::Round($License.RemainingGracePeriod / 1440, 0) } else { "N/A" }
$RaportObjektid.Add([PSCustomObject]@{Kontroll = "Litsentsi jääk (päeva)"; Staatus = $DaysLeft})

# --- HTML LOOMINE ---
$Header = @"
<style>
    body { font-family: 'Segoe UI', Arial; margin: 30px; background-color: #f0f2f5; }
    table { border-collapse: collapse; width: 100%; max-width: 800px; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    th, td { padding: 12px 15px; border: 1px solid #e0e0e0; text-align: left; }
    th { background-color: #005a9e; color: white; text-transform: uppercase; font-size: 14px; }
    .TEHTUD { color: #28a745; font-weight: bold; }
    .TEGEMATA { color: #dc3545; font-weight: bold; }
    h2 { color: #1a1a1a; }
</style>
"@

$HtmlBody = $RaportObjektid | ConvertTo-Html -Fragment | Out-String
$HtmlBody = $HtmlBody -replace "<td>TEHTUD</td>", "<td class='TEHTUD'>TEHTUD</td>"
$HtmlBody = $HtmlBody -replace "<td>TEGEMATA</td>", "<td class='TEGEMATA'>TEGEMATA</td>"

$FinalHTML = ConvertTo-Html -Head $Header -Body "<h2>Serveri audit: $MyDomain ($(Get-Date -Format 'dd.MM.yyyy HH:mm'))</h2> $HtmlBody"
$Path = "$home\Desktop\Raport.html"
$FinalHTML | Out-File $Path -Encoding utf8

Write-Host "Raport on valmis: $Path" -ForegroundColor Cyan
Invoke-Item $Path