<# =====================================================================
 HyperV-ClusterHealth.ps1
 - Cluster + VM inventory, HTML + Excel export
 - Runs in PS 7.x (uses WinPSCompat for Hyper-V/FailoverClusters) or PS 5.1
 - Logging: transcript to run.log in the report folder
 - Guest OS:
      1) PS Direct on OWNER HOST (if -GuestCredential provided AND host WinRM reachable) using -VMId
      2) KVP (Data Exchange) via Msvm_KvpExchangeComponent using VM GUID
 - Windows 11 normalization when build ≥ 22000
 - Optional -ForceWinPS to auto relaunch in Windows PowerShell 5.1
 ===================================================================== #>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ClusterName,

    [Parameter()]
    [string]$ReportRoot = (Join-Path -Path $PSScriptRoot -ChildPath "..\reports"),

    [Parameter()]
    [System.Management.Automation.PSCredential]$GuestCredential,

    [switch]$ForceWinPS
)

$ErrorActionPreference = 'Stop'

# --- Optional relaunch in Windows PowerShell 5.1 for a no-compat run ---
if ($ForceWinPS -and $PSVersionTable.PSEdition -eq 'Core') {
    # Avoid relaunch loops
    if (-not (Get-Variable -Name __LaunchedWinPS -Scope Global -ErrorAction SilentlyContinue)) {
        $argsList = @('-NoProfile','-ExecutionPolicy','Bypass','-File', $MyInvocation.MyCommand.Path, '-ClusterName', $ClusterName)
        if ($ReportRoot)       { $argsList += @('-ReportRoot', $ReportRoot) }
        if ($GuestCredential)  { $argsList += @('-GuestCredential', '$cred') }
        # If creds were provided, forward via temp variable in WinPS
        if ($GuestCredential) {
            $credExportPath = Join-Path $env:TEMP "hvcred.xml"
            try { $GuestCredential | Export-Clixml -Path $credExportPath -Force }
            catch { Write-Warning ("Failed to export credential: {0}" -f $_.Exception.Message) }
            $ps = Start-Process -FilePath "powershell.exe" -ArgumentList @(
                '-NoProfile','-ExecutionPolicy','Bypass','-Command',
                @"
`$global:__LaunchedWinPS = `$true;
`$cred = Import-Clixml -Path '$credExportPath';
& '$($MyInvocation.MyCommand.Path)' -ClusterName '$ClusterName' -ReportRoot '$ReportRoot' -GuestCredential `$cred
"@
            ) -PassThru -WindowStyle Normal
            exit
        } else {
            Start-Process -FilePath "powershell.exe" -ArgumentList $argsList -WindowStyle Normal | Out-Null
            exit
        }
    }
}

# --- PS7 compatibility import for Windows-only modules ---
function Import-CompatModule {
    param([Parameter(Mandatory)][string]$Name,[switch]$Required)
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            if ((Get-Command Import-Module).Parameters.ContainsKey('UseWindowsPowerShell')) {
                Import-Module -Name $Name -UseWindowsPowerShell -Force -ErrorAction Stop
            } else {
                if (-not (Get-PSSession -Name WinPSCompat -ErrorAction SilentlyContinue)) {
                    $script:WinPSCompat = New-PSSession -Name WinPSCompat -ConfigurationName Microsoft.PowerShell
                } else { $script:WinPSCompat = Get-PSSession -Name WinPSCompat }
                Import-Module -Name $Name -PSSession $script:WinPSCompat -Global -Force -ErrorAction Stop
            }
        } else {
            Import-Module -Name $Name -Force -ErrorAction Stop
        }
    } catch {
        $m = "Failed to import module '$Name': $($_.Exception.Message)"
        if ($Required) { throw $m } else { Write-Warning $m }
    }
}
Import-CompatModule -Name FailoverClusters -Required
Import-CompatModule -Name Hyper-V -Required

# --- ImportExcel (optional for XLSX) ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try { Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop }
    catch { Write-Warning "ImportExcel not installed: $($_.Exception.Message). Excel export will be skipped." }
}
Import-Module ImportExcel -ErrorAction SilentlyContinue | Out-Null

# --- Output directories + transcript ---
try {
    if (-not (Test-Path $ReportRoot)) { New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null }
} catch {
    $ReportRoot = Join-Path (Get-Location) "reports"
    if (-not (Test-Path $ReportRoot)) { New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null }
}

$stamp     = (Get-Date).ToString("yyyyMMdd-HHmmss")
$ReportDir = Join-Path $ReportRoot $stamp
New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null

$excelPath = Join-Path $ReportDir "ClusterHealth.xlsx"
$htmlPath  = Join-Path $ReportDir "ClusterHealth.html"
$logPath   = Join-Path $ReportDir "run.log"

Start-Transcript -Path $logPath -Force | Out-Null

Write-Host "Created report directory: $ReportDir"
Write-Host "[+] Connecting to cluster: $ClusterName"
$cluster = $null
try { $cluster = Get-Cluster -Name $ClusterName -ErrorAction Stop } catch { Write-Warning ("Get-Cluster failed: {0}" -f $_.Exception.Message) }

# --- Helpers ---
function Get-SafeUptime {
    param([string]$Dmtf,[string]$ComputerName)
    if ($Dmtf) {
        try {
            $dt = [System.Management.ManagementDateTimeConverter]::ToDateTime($Dmtf)
            if ($dt -and $dt -lt (Get-Date).AddYears(10)) { return ((Get-Date) - $dt) }
        } catch { }
    }
    try {
        $perf = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_PerfFormattedData_PerfOS_System -ErrorAction Stop
        if ($perf.SystemUpTime -ge 0) { return [TimeSpan]::FromSeconds([double]$perf.SystemUpTime) }
    } catch { }
    return $null
}

function Normalize-WindowsClientName {
    param([string]$ProductName,[int]$Build,[string]$DisplayVersion)
    if ($ProductName -match 'Windows Server') {
        if ($DisplayVersion) { return "$ProductName ($DisplayVersion, build $Build)" }
        else { return "$ProductName (build $Build)" }
    }
    $name = $ProductName
    if ($Build -ge 22000) {
        if ($name -match 'Windows 11') { }
        elseif ($name -match 'Windows 10') { $name = $name -replace 'Windows 10','Windows 11' }
        elseif ($name) { $name = "Windows 11 ($name)" }
        else { $name = 'Windows 11' }
    } elseif (-not $name) {
        $name = 'Windows (unknown edition)'
    }
    if ($DisplayVersion) { return "$name ($DisplayVersion, build $Build)" }
    elseif ($Build)      { return "$name (build $Build)" }
    else                 { return $name }
}

function Encode-Html { param([string]$Text)
    if ($null -eq $Text) { return "" }
    ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'",'&#39;')
}

function Test-HostRemoting {
    param([string]$ComputerName)
    try { return [bool](Test-WSMan -ComputerName $ComputerName -ErrorAction Stop) } catch { return $false }
}

# --- Diagnostics (hosts) ---
function Test-HostPrereqs {
    param([string[]]$Hosts)
    foreach ($hn in $Hosts) {
        $rem = Test-HostRemoting -ComputerName $hn
        if (-not $rem) { Write-Warning ("[Diagnostics] Host {0}: WinRM not reachable (PS Direct will be skipped)" -f $hn) }
        try {
            Get-CimInstance -ComputerName $hn -Namespace 'root/virtualization/v2' -ClassName Msvm_ComputerSystem -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning ("[Diagnostics] Host {0}: cannot query 'root/virtualization/v2' (Hyper-V WMI). {1}" -f $hn, $_.Exception.Message)
        }
    }
}

# Fixed columns
$NodeColumns = @('NodeName','Status','OS','OSVersion','Uptime','Manufacturer','Model','CPUModel','Cores','LogicalProcs','TotalMemoryGB','DiskSummary')
$VMColumns   = @('VMName','HostName','State','Uptime','CPUPercent','MemoryAssignedMB','Generation','ConfigurationVer','GuestOS','VHDUsage')

# --- Node collection ---
$nodeResults = New-Object System.Collections.Generic.List[object]
$clusterNodes = @()
try { $clusterNodes = Get-ClusterNode -Cluster $ClusterName } catch { throw "Get-ClusterNode failed: $($_.Exception.Message)" }

# run diagnostics early
Test-HostPrereqs -Hosts ($clusterNodes | ForEach-Object { $_.Name })

foreach ($n in $clusterNodes) {
    $nodeName = $n.Name
    Write-Host "[>] Gathering data for node: $nodeName"
    try {
        $cs=$null;$cpu=$null;$os=$null;$ld=$null
        try { $cs = Get-CimInstance -ComputerName $nodeName -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { }
        try { $cpu= Get-CimInstance -ComputerName $nodeName -ClassName Win32_Processor -ErrorAction Stop } catch { }
        try { $os = Get-CimInstance -ComputerName $nodeName -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { }
        try { $ld = Get-CimInstance -ComputerName $nodeName -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop } catch { }

        $uptime = Get-SafeUptime -Dmtf $os.LastBootUpTime -ComputerName $nodeName

        $cpuName   = ($cpu | Select-Object -First 1).Name
        $cores     = ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum
        $lps       = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        $memGB     = if ($cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory/1GB,2) } else { $null }

        $diskSummary = if ($ld) {
            ($ld | Sort-Object DeviceID | ForEach-Object {
                $free = if ($_.FreeSpace) { [math]::Round($_.FreeSpace/1GB,1) } else { $null }
                $size = if ($_.Size)      { [math]::Round($_.Size/1GB,1) }      else { $null }
                if ($free -ne $null -and $size -ne $null) { "{0}: {1} GB free of {2} GB" -f $_.DeviceID, $free, $size } else { "{0}: n/a" -f $_.DeviceID }
            }) -join "; "
        } else { $null }

        $nodeResults.Add([PSCustomObject]@{
            NodeName        = $nodeName
            Status          = $n.State
            OS              = $os.Caption
            OSVersion       = $os.Version
            Uptime          = if ($uptime) { "{0:%d}d {0:hh}h {0:mm}m" -f $uptime } else { $null }
            Manufacturer    = $cs.Manufacturer
            Model           = $cs.Model
            CPUModel        = $cpuName
            Cores           = $cores
            LogicalProcs    = $lps
            TotalMemoryGB   = $memGB
            DiskSummary     = $diskSummary
        }) | Out-Null
    } catch {
        Write-Warning ("Failed to gather data for node {0}: {1}" -f $nodeName, $_.Exception.Message)
    }
}

# --- Build VM descriptors by reading VMId from cluster resources ---
$vmDescriptors = @()
try {
    $allResources = Get-ClusterResource -Cluster $ClusterName
    $vmGroups = Get-ClusterGroup -Cluster $ClusterName | Where-Object { $_.GroupType -eq 'VirtualMachine' }
    foreach ($g in $vmGroups) {
        $vmRes = $allResources | Where-Object { $_.OwnerGroup -eq $g -and $_.ResourceType -eq 'Virtual Machine' } | Select-Object -First 1
        if ($vmRes) {
            $vmIdParam = $null
            try { $vmIdParam = $vmRes | Get-ClusterParameter -Name 'VMId' -ErrorAction Stop } catch { }
            if (-not $vmIdParam) {
                try { $vmIdParam = $vmRes | Get-ClusterParameter -Name 'VirtualMachineId' -ErrorAction Stop } catch { }
            }
            $vmId = $null
            if ($vmIdParam -and $vmIdParam.Value) {
                try { $vmId = [Guid]$vmIdParam.Value } catch { }
            }
            $owner = $null; try { $owner = $g.OwnerNode.Name } catch { }
            $vmDescriptors += [PSCustomObject]@{
                ClusterGroup = $g.Name
                VMId         = $vmId
                OwnerNode    = $owner
            }
        }
    }
} catch {
    Write-Warning ("Failed to read VM resources: {0}" -f $_.Exception.Message)
}

# --- VM helpers ---
$candidateHosts = @($clusterNodes | ForEach-Object { $_.Name })

function Resolve-VMById {
    param([Guid]$VmId,[string]$OwnerNode,[string[]]$AllHosts)
    $vm = $null
    if ($VmId -ne [Guid]::Empty) {
        if ($OwnerNode) {
            try { $vm = Get-VM -ComputerName $OwnerNode -Id $VmId -ErrorAction Stop } catch { }
            if ($vm) { return ,@($vm,$OwnerNode) }
        }
        foreach ($hv in $AllHosts) {
            try { $vm = Get-VM -ComputerName $hv -Id $VmId -ErrorAction Stop } catch { }
            if ($vm) { return ,@($vm,$hv) }
        }
    }
    return ,@( $null, $null )
}

# 1) PS Direct (only if host remoting reachable) using VMId
function Get-GuestOs-PSDirect {
    param([Guid]$VmId,[string]$HostName,[System.Management.Automation.PSCredential]$Cred)
    if (-not $Cred) { return $null }
    if (-not (Test-HostRemoting -ComputerName $HostName)) {
        Write-Warning ("PS Direct skipped (host {0} WinRM not reachable)" -f $HostName)
        return $null
    }
    try {
        $result = Invoke-Command -ComputerName $HostName -ErrorAction Stop -ScriptBlock {
            param($id)
            try {
                $inner = {
                    try {
                        Set-Service -Name vmickvpexchange -StartupType Automatic -ErrorAction SilentlyContinue
                        Start-Service -Name vmickvpexchange -ErrorAction SilentlyContinue
                        $cv = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
                        $kv = Get-ItemProperty -Path $cv -ErrorAction Stop
                        $prod  = [string]$kv.ProductName
                        $disp  = [string]$kv.DisplayVersion
                        $build = 0
                        if ($kv.CurrentBuild) { [int]::TryParse([string]$kv.CurrentBuild,[ref]$build) | Out-Null }
                        if ($prod -match 'Windows Server') {
                            if ($disp) { return "$prod ($disp, build $build)" }
                            else       { return "$prod (build $build)" }
                        } else {
                            $name = $prod
                            if ($build -ge 22000) {
                                if ($name -match 'Windows 11') { }
                                elseif ($name -match 'Windows 10') { $name = $name -replace 'Windows 10','Windows 11' }
                                elseif ($name) { $name = "Windows 11 ($name)" }
                                else { $name = 'Windows 11' }
                            } elseif (-not $name) {
                                $name = 'Windows (unknown edition)'
                            }
                            if ($disp) { return "$name ($disp, build $build)" }
                            elseif ($build) { return "$name (build $build)" }
                            else { return $name }
                        }
                    } catch { "Unavailable (guest registry not accessible)" }
                }
                Invoke-Command -VMId $id -Credential $using:Cred -ScriptBlock $inner -ErrorAction Stop
            } catch { "Unavailable (PS Direct failed on host)" }
        } -ArgumentList $VmId
        return ($result | Select-Object -First 1)
    } catch {
        Write-Warning ("PS Direct failed on host {0}: {1}" -f $HostName, $_.Exception.Message)
        return $null
    }
}

# 2) KVP via VM GUID with expanded parsing (PS5.1-safe)
function Get-GuestOs-Kvp {
    param([Guid]$VmId,[string]$HostName)
    try {
        Invoke-Command -ComputerName $HostName -ErrorAction Stop -ScriptBlock {
            param($id)
            try {
                $idStr = $id.ToString()
                $cs = Get-CimInstance -Namespace 'root/virtualization/v2' -ClassName Msvm_ComputerSystem -Filter ("Name='{0}'" -f $idStr) -ErrorAction SilentlyContinue
                if (-not $cs) { return $null }

                $kvp = Get-CimAssociatedInstance -InputObject $cs -Association Msvm_SystemDevice -ErrorAction Stop |
                       Where-Object { $_.CimClass.CimClassName -eq 'Msvm_KvpExchangeComponent' } | Select-Object -First 1
                if (-not $kvp) { return $null }

                $items = @()
                if ($kvp.GuestIntrinsicExchangeItems) { $items += $kvp.GuestIntrinsicExchangeItems }
                if ($kvp.GuestExchangeItems)         { $items += $kvp.GuestExchangeItems }
                if (-not $items) { return $null }

                $map = @{}
                foreach ($str in $items) {
                    try {
                        $xml   = [xml]$str
                        $nm    = ($xml.INSTANCE.PROPERTY | Where-Object NAME -eq 'Name').VALUE
                        $val   = ($xml.INSTANCE.PROPERTY | Where-Object NAME -eq 'Data').VALUE
                        if ($nm) { $map[$nm] = $val }
                    } catch { }
                }

                $osName    = $map['OSName']
                $osVersion = $map['OSVersion']   # e.g. 10.0.22631
                $build     = $null
                if ($osVersion) {
                    $parts = $osVersion -split '\.'
                    if ($parts.Length -ge 3) { [int]::TryParse($parts[2],[ref]$build) | Out-Null }
                }
                if (-not $build -and $map.ContainsKey('BuildNumber')) {
                    [int]::TryParse([string]$map['BuildNumber'],[ref]$build) | Out-Null
                }

                $disp = $null
                if ($map.ContainsKey('DisplayVersion') -and $map['DisplayVersion']) { $disp = $map['DisplayVersion'] }
                elseif ($map.ContainsKey('ReleaseId') -and $map['ReleaseId']) { $disp = $map['ReleaseId'] }

                if ($osName -match 'Windows Server') {
                    if ($disp) { return "$osName ($disp, build $build)" }
                    elseif ($build) { return "$osName (build $build)" }
                    else { return $osName }
                } elseif ($osName -match 'Windows') {
                    if ($build -ge 22000 -and $osName -match 'Windows 10') { $osName = $osName -replace 'Windows 10','Windows 11' }
                    if ($disp) { return "$osName ($disp, build $build)" }
                    elseif ($build) { return "$osName (build $build)" }
                    else { return $osName }
                } else {
                    return $osName
                }
            } catch { return $null }
        } -ArgumentList $VmId
    } catch {
        Write-Warning ("KVP read failed on host {0}: {1}" -f $HostName, $_.Exception.Message)
        return $null
    }
}

function Get-VhdUsageSummary {
    param([string]$HostName,[string[]]$VhdPaths)
    $summ = @()
    foreach ($p in $VhdPaths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        try {
            $v = Get-VHD -ComputerName $HostName -Path $p -ErrorAction Stop
            $vs = [double]$v.VirtualSize
            $fs = [double]$v.FileSize
            $vsText = if ($vs -gt 0) { "{0:N1} GB" -f ($vs/1GB) } else { "n/a" }
            $fsText = if ($fs -ge 0) { "{0:N1} GB" -f ($fs/1GB) } else { "n/a" }
            $leaf   = Split-Path $p -Leaf
            $par    = if ($v.ParentPath) { "(parent: {0})" -f (Split-Path $v.ParentPath -Leaf) } else { "" }
            $summ  += "{0}: {1} of {2} {3}" -f $leaf, $fsText, $vsText, $par
        } catch {
            $summ += "{0}: n/a" -f (Split-Path $p -Leaf)
        }
    }
    return ($summ -join "; ")
}

# --- Gather VMs using descriptors (VMId + owner host) ---
$vmResults = New-Object System.Collections.Generic.List[object]
foreach ($d in $vmDescriptors) {
    $vmId = $d.VMId
    $vmNameForDisplay = $d.ClusterGroup
    Write-Host "[>] Gathering VM: $vmNameForDisplay"

    $resolved = Resolve-VMById -VmId $vmId -OwnerNode $d.OwnerNode -AllHosts $candidateHosts
    $vm        = $resolved[0]
    $ownerHost = $resolved[1]
    if (-not $vm -or -not $ownerHost) {
        Write-Warning ("Skipping VM {0}: no reachable host found." -f $vmNameForDisplay)
        continue
    }

    # Ensure host-side KVP is enabled (best effort)
    try {
        $kvpSvc = Get-VMIntegrationService -ComputerName $ownerHost -VMId $vm.Id -Name 'Key-Value Pair Exchange' -ErrorAction Stop
        if (-not $kvpSvc.Enabled) {
            Enable-VMIntegrationService -ComputerName $ownerHost -VMId $vm.Id -Name 'Key-Value Pair Exchange' -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    # Guest OS: PS Direct (if host remoting OK & creds) else KVP
    $guestOs = $null
    if ($GuestCredential -and ($vm.State -eq 'Running')) {
        $guestOs = Get-GuestOs-PSDirect -VmId $vm.Id -HostName $ownerHost -Cred $GuestCredential
        if ($guestOs -and $guestOs -is [array]) { $guestOs = $guestOs | Select-Object -First 1 }
    }
    if (-not $guestOs) {
        $guestOs = Get-GuestOs-Kvp -VmId $vm.Id -HostName $ownerHost
    }
    if (-not $guestOs) {
        $guestOs = "Unavailable (enable Data Exchange in guest or ensure host WinRM + provide -GuestCredential)"
    }

    # VHD usage
    $vhdPaths = @()
    try {
        $vhdPaths = (Get-VMHardDiskDrive -ComputerName $ownerHost -VMId $vm.Id -ErrorAction Stop | Select-Object -ExpandProperty Path) | Where-Object { $_ }
    } catch { }
    $vhdSummary = if ($vhdPaths.Count -gt 0) { Get-VhdUsageSummary -HostName $ownerHost -VhdPaths $vhdPaths } else { $null }

    # Convert MemoryAssigned bytes -> MB; ensure ConfigVer is string
    $memMB     = if ($vm.MemoryAssigned -ge 0) { [int]([math]::Round($vm.MemoryAssigned/1MB,0)) } else { $null }
    $configVer = [string]$vm.Version

    $vmResults.Add([PSCustomObject]@{
        VMName           = $vm.Name
        HostName         = $ownerHost
        State            = $vm.State
        Uptime           = $vm.Uptime
        CPUPercent       = $vm.CPUUsage
        MemoryAssignedMB = $memMB
        Generation       = $vm.Generation
        ConfigurationVer = $configVer
        GuestOS          = $guestOs
        VHDUsage         = $vhdSummary
    }) | Out-Null
}

# --- Excel export (if available) ---
if (Get-Module -Name ImportExcel -ListAvailable) {
    try {
        $nodeResults | Select-Object -Property $NodeColumns | Export-Excel -Path $excelPath -WorksheetName "ClusterNodes" -AutoSize -BoldTopRow -FreezeTopRow
        $vmResults   | Select-Object -Property $VMColumns   | Export-Excel -Path $excelPath -WorksheetName "ClusterVMs"   -AutoSize -BoldTopRow -FreezeTopRow
        Write-Host "Excel written: $excelPath"
    } catch { Write-Warning ("Excel export failed: {0}" -f $_.Exception.Message) }
} else {
    Write-Warning "ImportExcel not found; skipping Excel export."
}

# --- HTML export (DataTables) ---
function ConvertTo-HtmlTable {
    param([Parameter(Mandatory)][System.Collections.IEnumerable]$Objects,[Parameter(Mandatory)][string]$TableId,[Parameter(Mandatory)][string[]]$Columns)
    $html = ""
    $html += "<table id=""$TableId"" class=""display compact stripe"" style=""width:100%"">`n<thead><tr>"
    foreach ($p in $Columns) { $html += "<th>$p</th>" }
    $html += "</tr></thead>`n<tbody>`n"
    foreach ($o in $Objects) {
        $html += "<tr>"
        foreach ($p in $Columns) {
            $val = $o.$p
            if ($val -is [TimeSpan]) { $val = "{0:%d}d {0:hh}h {0:mm}m" -f $val }
            $html += "<td>{0}</td>" -f (Encode-Html ([string]$val))
        }
        $html += "</tr>`n"
    }
    $html += "</tbody></table>`n"
    return $html
}

$nodesTable = ConvertTo-HtmlTable -Objects $nodeResults -TableId "nodes" -Columns $NodeColumns
$vmsTable   = ConvertTo-HtmlTable -Objects $vmResults   -TableId "vms"   -Columns $VMColumns

$header = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Hyper-V Cluster Health Report</title>
<link rel="preconnect" href="https://cdn.datatables.net" />
<link rel="preconnect" href="https://code.jquery.com" />
<link href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css" rel="stylesheet" />
<style>
:root { --ink:#0b5cab; --bg:#f6f7fb; --fg:#1f2937; --muted:#6b7280; --card:#ffffff; --line:#e2e6ef; }
html,body{margin:0;padding:0;background:var(--bg);color:var(--fg);font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;}
.wrap{max-width:1200px;margin:32px auto;padding:0 20px;}
h1{color:var(--ink);margin:0 0 6px;}
h2{color:var(--ink);margin:18px 0 10px;}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 16px;margin:16px 0;}
.meta{color:var(--muted);font-size:13px;margin-bottom:10px}
table.dataTable thead th{background:var(--ink);color:#fff}
td.wrap{word-break:break-word;max-width:0}
footer{color:var(--muted);margin:36px 0 48px;text-align:center}
.small{color:var(--muted);font-size:12px}
</style>
</head>
<body>
<div class="wrap">
'@

$intro = @"
<h1>Hyper-V Cluster Health — $ClusterName</h1>
<div class=""meta"">Generated: $(Get-Date)</div>

<section>
  <h2>Cluster Nodes</h2>
  <div class=""small"">Search, sort, and paginate. Columns are fixed for consistency.</div>
  $nodesTable
</section>

<section>
  <h2>Virtual Machines</h2>
  <div class=""small"">Guest OS via PS Direct (host WinRM + creds) or KVP fallback (GUID-based); Windows 11 normalized if build ≥ 22000.</div>
  $vmsTable
</section>
"@

$footer = @'
<footer>© Cluster Health · PowerShell · DataTables</footer>
</div>
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>
<script>
$(document).ready(function(){
  $('#nodes').DataTable({
    pageLength: 25,
    order: [[0,'asc']],
    autoWidth: false,
    scrollX: true,
    columnDefs: [
      { targets: [8,9,10], className: 'dt-right', type: 'num' },
      { targets: [11], className: 'wrap' }
    ]
  });

  // VMs: [0]=VMName,1=Host,2=State,3=Uptime,4=CPU%,5=MemMB,6=Gen,7=ConfigVer,8=GuestOS,9=VHD
  $('#vms').DataTable({
    pageLength: 25,
    order: [[0,'asc']],
    autoWidth: false,
    scrollX: true,
    columnDefs: [
      { targets: [4,5,6,7], className: 'dt-right', type: 'num' },
      { targets: [0,8,9], className: 'wrap' }
    ]
  });
});
</script>
</body>
</html>
'@

Set-Content -Path $htmlPath -Value ($header + $intro + $footer) -Encoding UTF8
Write-Host "HTML written: $htmlPath"
Write-Host "[✔] Reports generated under: $ReportDir"

Stop-Transcript | Out-Null
