<#
.SYNOPSIS
  Hyper-V Cluster Health Report (HTML + Excel) with guest OS fixes (KVP + PowerShell Direct)
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ClusterName,
    [Parameter()]
    [System.Management.Automation.PSCredential]$GuestCredential
)

function Import-ModuleSafe {
    param([string]$Name, [switch]$SkipEditionCheck)
    try {
        if ($SkipEditionCheck) { Import-Module $Name -SkipEditionCheck -ErrorAction Stop }
        else { Import-Module $Name -ErrorAction Stop }
    } catch {
        if ($PSVersionTable.PSEdition -eq 'Core') { Import-Module $Name -UseWindowsPowerShell -ErrorAction Stop }
        else { throw }
    }
}

Import-ModuleSafe -Name FailoverClusters -SkipEditionCheck
Import-ModuleSafe -Name Hyper-V

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop }
    catch { Write-Warning ("Failed to install ImportExcel automatically: {0}" -f $_.Exception.Message) }
}
Import-Module ImportExcel -ErrorAction Stop

$baseDir   = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$reportDir = Join-Path $baseDir "reports\$timestamp"
$htmlPath  = Join-Path $reportDir "ClusterHealth.html"
$xlsxPath  = Join-Path $reportDir "ClusterHealth.xlsx"

if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
Write-Host ("Created report directory: {0}" -f $reportDir)
Write-Host ("[+] Connecting to cluster: {0}" -f $ClusterName) -ForegroundColor Cyan

function Ensure-DataExchangeEnabled {
    param([string]$VmName, [string]$HostName)
    try {
        $svc = Get-VMIntegrationService -ComputerName $HostName -VMName $VmName -Name "Key-Value Pair Exchange" -ErrorAction Stop
        if (-not $svc.Enabled) {
            Enable-VMIntegrationService -ComputerName $HostName -VMName $VmName -Name "Key-Value Pair Exchange" -ErrorAction Stop
            Write-Host ("Enabled Data Exchange for VM '{0}' on host '{1}'" -f $VmName, $HostName)
        }
        return $true
    } catch {
        Write-Warning ("Could not enable Data Exchange for VM {0} on {1}: {2}" -f $VmName, $HostName, $_.Exception.Message)
        return $false
    }
}

function Try-StartGuestKVP {
    param([string]$VmName, [string]$HostName, [System.Management.Automation.PSCredential]$Credential)
    if (-not $Credential) { return $false }
    try {
        Invoke-Command -ComputerName $HostName -ErrorAction Stop -ScriptBlock {
            param($InnerVmName, $Cred)
            Invoke-Command -VMName $InnerVmName -Credential $Cred -ErrorAction Stop -ScriptBlock {
                $svc = Get-Service -Name vmickvpexchange -ErrorAction SilentlyContinue
                if (-not $svc) { throw "vmickvpexchange not found (non-Windows guest or components missing)" }
                Set-Service -Name vmickvpexchange -StartupType Automatic -ErrorAction Stop
                if ($svc.Status -ne 'Running') { Start-Service -Name vmickvpexchange -ErrorAction Stop }
            }
        } -ArgumentList $VmName, $Credential
        Write-Host ("Started/ensured guest KVP service for VM '{0}' via host '{1}'" -f $VmName, $HostName)
        return $true
    } catch {
        Write-Warning ("PS Direct guest KVP start failed for VM {0} on host {1}: {2}" -f $VmName, $HostName, $_.Exception.Message)
        return $false
    }
}

function Get-GuestOSViaPSDirect {
    param([string]$VmName, [string]$HostName, [System.Management.Automation.PSCredential]$Credential)
    if (-not $Credential) { return $null }
    try {
        $osStr = Invoke-Command -ComputerName $HostName -ErrorAction Stop -ScriptBlock {
            param($InnerVmName, $Cred)
            Invoke-Command -VMName $InnerVmName -Credential $Cred -ErrorAction Stop -ScriptBlock {
                try {
                    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
                    $prod=$cv.ProductName; $disp=$cv.DisplayVersion; $build=$cv.CurrentBuild; $ubr=$cv.UBR
                    if (-not $prod) { $prod = "Windows (unknown edition)" }
                    if ($disp) { "{0} {1} (build {2}.{3})" -f $prod, $disp, $build, $ubr }
                    else { "{0} (build {1}.{2})" -f $prod, $build, $ubr }
                } catch {
                    $wmi = Get-CimInstance Win32_OperatingSystem
                    if ($wmi.Caption) { $wmi.Caption + " " + $wmi.Version } else { $null }
                }
            }
        } -ArgumentList $VmName, $Credential | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($osStr)) { return $null }
        return $osStr
    } catch { return $null }
}

function Get-ClusterVMResources { param([string]$ClusterName)
    Get-ClusterResource -Cluster $ClusterName | Where-Object { $_.ResourceType -eq 'Virtual Machine' }
}
function Get-ClusterVMId { param($ClusterResource)
    $vmId=$null
    try{$p=Get-ClusterParameter -InputObject $ClusterResource -Name VmId -ErrorAction SilentlyContinue; if($p.Value){return $p.Value}}catch{}
    try{$p=Get-ClusterParameter -InputObject $ClusterResource -Name VirtualMachineId -ErrorAction SilentlyContinue; if($p.Value){return $p.Value}}catch{}
    return $vmId
}
function Resolve-VMFromResource {
    param($ClusterResource,[string[]]$CandidateHosts)
    $vmId=Get-ClusterVMId -ClusterResource $ClusterResource
    $nameFromGroup=$null; try{$nameFromGroup=$ClusterResource.OwnerGroup.Name}catch{}
    $nameFromRes=($ClusterResource.Name -replace '^Virtual Machine\s*','').Trim()
    foreach($h in $CandidateHosts){
        try{ if($vmId){ return Get-VM -Id $vmId -ComputerName $h -ErrorAction Stop } }catch{}
        try{ if($nameFromGroup){ return Get-VM -Name $nameFromGroup -ComputerName $h -ErrorAction Stop } }catch{}
        try{ if($nameFromRes){ return Get-VM -Name $nameFromRes -ComputerName $h -ErrorAction Stop } }catch{}
    }
    return $null
}

$nodes = Get-ClusterNode -Cluster $ClusterName
$allNodeNames = $nodes | Select-Object -ExpandProperty Name

$nodeResults = foreach ($node in $nodes) {
    $n = $node.Name
    Write-Host ("[>] Gathering data for node: {0}" -f $n)
    try {
        $os    = Get-CimInstance Win32_OperatingSystem      -ComputerName $n
        $cs    = Get-CimInstance Win32_ComputerSystem       -ComputerName $n
        $cpu   = Get-CimInstance Win32_Processor            -ComputerName $n
        $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $n

        $uptime = "Unavailable"
        if ($os.LastBootUpTime -and $os.LastBootUpTime.Length -ge 8) {
            try {
                $bt=[Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
                $span=(Get-Date) - $bt
                $uptime="{0:dd}d {0:hh}h {0:mm}m" -f $span
            } catch {
                Write-Warning ("Uptime conversion failed for node {0}: {1}" -f $n, $_.Exception.Message)
            }
        }

        $diskSummary = ($disks | ForEach-Object {
            "{0}: {1:N1}GB free of {2:N1}GB" -f $_.DeviceID, ($_.FreeSpace/1GB), ($_.Size/1GB)
        }) -join "; "

        [PSCustomObject]@{
            NodeName          = $n
            Status            = $node.State
            OS                = $os.Caption
            OSVersion         = $os.Version
            Uptime            = $uptime
            Manufacturer      = $cs.Manufacturer
            Model             = $cs.Model
            TotalMemoryGB     = "{0:N1}" -f ($cs.TotalPhysicalMemory/1GB)
            CPUModel          = $cpu.Name
            CPUCores          = ($cpu | Measure-Object NumberOfCores             -Sum).Sum
            LogicalProcessors = ($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
            DiskSummary       = $diskSummary
        }
    } catch {
        Write-Warning ("Failed to gather data for node {0}: {1}" -f $n, $_.Exception.Message)
    }
}

$vmResults = @()
$vmResources = Get-ClusterVMResources -ClusterName $ClusterName

foreach ($vmRes in $vmResources) {
    $primary=$null; try{$primary=$vmRes.OwnerNode.Name}catch{$primary=$null}
    $candidates = if ($primary) { @($primary) } else { $allNodeNames }
    $displayName = $vmRes.Name

    $vm = Resolve-VMFromResource -ClusterResource $vmRes -CandidateHosts $candidates
    if (-not $vm) { Write-Warning ("Skipping VM {0}: no reachable host found." -f $displayName); continue }

    Write-Host ("[>] Gathering VM: {0}" -f $vm.Name)

    $hostSideOk = Ensure-DataExchangeEnabled -VmName $vm.Name -HostName $vm.ComputerName

    $didStart = $false
    if ($hostSideOk -and $vm.State -eq 'Running' -and $GuestCredential) {
        $didStart = Try-StartGuestKVP -VmName $vm.Name -HostName $vm.ComputerName -Credential $GuestCredential
        if ($didStart) { Start-Sleep -Seconds 6 }
    }

    $guestOS = $vm.GuestOperatingSystem
    $guestOSSource = if ($guestOS) { 'KVP' } else { 'None' }
    if ([string]::IsNullOrWhiteSpace($guestOS) -and $vm.State -eq 'Running' -and $GuestCredential) {
        $osDirect = Get-GuestOSViaPSDirect -VmName $vm.Name -HostName $vm.ComputerName -Credential $GuestCredential
        if ($osDirect) { $guestOS=$osDirect; $guestOSSource='PSDirect' }
    }
    if (-not $guestOS) { $guestOS = "Unavailable (enable Data Exchange in guest or supply -GuestCredential / accept prompt)" }

    $vhdSummary = "Unavailable"
    try {
        $vhds = Get-VMHardDiskDrive -VMName $vm.Name -ComputerName $vm.ComputerName -ErrorAction Stop
        $summaryList = @()
        foreach ($vhd in $vhds) {
            $vhdPath = $vhd.Path
            try {
                $info = Get-VHD -Path $vhdPath -ComputerName $vm.ComputerName -ErrorAction Stop
                $parentNote = ""
                if ($info.VhdType -eq 'Differencing' -and $info.ParentPath) {
                    $parentNote = " (parent: {0})" -f ([IO.Path]::GetFileName($info.ParentPath))
                }
                $summaryList += "{0} ({1}): {2:N1}GB used of {3:N1}GB{4}" -f `
                    ([IO.Path]::GetFileName($vhdPath)), $info.VhdType, ($info.FileSize/1GB), ($info.Size/1GB), $parentNote
            } catch {
                Write-Warning ("VHD query failed for VM {0} on host {1} path {2}: {3}" -f `
                    $vm.Name, $vm.ComputerName, $vhdPath, $_.Exception.Message)
            }
        }
        if ($summaryList.Count -gt 0) { $vhdSummary = $summaryList -join "; " }
    } catch {
        Write-Warning ("Host VHD enumeration failed for VM {0}: {1}" -f $vm.Name, $_.Exception.Message)
    }

    $hostKvpStatus  = (Get-VMIntegrationService -ComputerName $vm.ComputerName -VMName $vm.Name -Name "Key-Value Pair Exchange" -ErrorAction SilentlyContinue |
                      Select-Object -First 1 -ExpandProperty Enabled)
    if ($null -eq $hostKvpStatus) { $hostKvpStatus = $false }

    $vmResults += [PSCustomObject]@{
        VMName              = $vm.Name
        HostName            = $vm.ComputerName
        State               = $vm.State
        Uptime              = $vm.Uptime.ToString("dd\.hh\:mm\:ss")
        CPUUsage            = $vm.CPUUsage
        MemoryAssignedMB    = [math]::Round($vm.MemoryAssigned/1MB,2)
        Generation          = $vm.Generation
        GuestOS             = $guestOS
        GuestOSSource       = $guestOSSource
        HostKVPEnabled      = $hostKvpStatus
        GuestKVPFixApplied  = $didStart
        VHDUsage            = $vhdSummary
        Version             = $vm.Version
    }
}

if ($nodeResults.Count -gt 0) { $nodeResults | Export-Excel -Path $xlsxPath -WorksheetName "ClusterNodes" -AutoSize }
else { [pscustomobject]@{ Info="No node data collected"} | Export-Excel -Path $xlsxPath -WorksheetName "ClusterNodes" -AutoSize }
if ($vmResults.Count -gt 0) { $vmResults | Export-Excel -Path $xlsxPath -WorksheetName "ClusterVMs" -AutoSize }
else { [pscustomobject]@{ Info="No VM data collected"} | Export-Excel -Path $xlsxPath -WorksheetName "ClusterVMs" -AutoSize }

$html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset='UTF-8'>
  <title>Hyper-V Cluster Health — $ClusterName</title>
  <style>
    body { font-family: Arial, Helvetica, sans-serif; background: #f6f7fb; padding: 24px; color: #222; }
    h1 { color: #0b5cab; margin-bottom: 0; }
    h2 { color: #0b5cab; margin-top: 28px; }
    .meta { color: #555; margin-bottom: 18px; }
    table { border-collapse: collapse; width: 100%; margin: 14px 0 28px; background: #fff; }
    th, td { border: 1px solid #e2e6ef; padding: 8px 10px; text-align: left; font-size: 13px; }
    th { background: #0b5cab; color: #fff; position: sticky; top: 0; }
    tr:nth-child(even) { background: #f2f6ff; }
    .muted { color: #6b7280; font-style: italic; }
  </style>
</head>
<body>
  <h1>Hyper-V Cluster Health Report</h1>
  <div class='meta'><strong>Cluster:</strong> $ClusterName &nbsp;&nbsp; <strong>Generated:</strong> $(Get-Date)</div>
  <h2>Cluster Nodes</h2>
"@

if ($nodeResults.Count -gt 0) {
    $html += "<table><tr>" +
        (($nodeResults[0].PSObject.Properties.Name | ForEach-Object { "<th>$_</th>" }) -join "") +
        "</tr>"
    foreach ($row in $nodeResults) {
        $html += "<tr>" + ($row.PSObject.Properties.Value | ForEach-Object { "<td>$_</td>" }) -join "" + "</tr>"
    }
    $html += "</table>"
} else { $html += "<p class='muted'>No node data collected.</p>" }

$html += "<h2>Virtual Machines</h2>"

if ($vmResults.Count -gt 0) {
    $html += "<table><tr>" +
        (($vmResults[0].PSObject.Properties.Name | ForEach-Object { "<th>$_</th>" }) -join "") +
        "</tr>"
    foreach ($row in $vmResults) {
        $html += "<tr>" + ($row.PSObject.Properties.Value | ForEach-Object { "<td>$_</td>" }) -join "" + "</tr>"
    }
    $html += "</table>"
} else { $html += "<p class='muted'>No VM data collected.</p>" }

$html += @"
</body>
</html>
"@

Set-Content -Path $htmlPath -Value $html -Encoding UTF8
Write-Host ("Excel nodes written: {0}" -f $xlsxPath)
Write-Host ("Excel VMs written:  {0}" -f $xlsxPath)
Write-Host ("HTML written: {0}" -f $htmlPath)
Write-Host ("`n[✔] All reports saved under: {0}" -f $reportDir) -ForegroundColor Green
