<!DOCTYPE html>
<html lang="en">
</head>
<body>
<div class="wrap">
<header>
  <h1>Hyper-V Cluster Health Report</h1>
  <p>Generate a clean HTML dashboard and Excel workbook for your Hyper-V Failover Cluster — nodes + VMs — with optional fixes for “guest OS unknown”.</p>
  <span class="badge">PowerShell</span>
  <span class="badge">Hyper-V</span>
  <span class="badge">Failover Clustering</span>
</header>

<section>
  <h2>Highlights</h2>
  <div class="grid">
    <div><strong>Cluster Nodes</strong><ul><li>OS/version, uptime, CPU model/cores/LPs</li><li>Memory and per-drive free space</li></ul></div>
    <div><strong>Virtual Machines</strong><ul><li>State, uptime, CPU %, memory assigned</li><li>Generation, config version, VHD usage</li><li>Guest OS via KVP, PS Direct fallback</li></ul></div>
    <div><strong>Exports</strong><ul><li>HTML (styled) + Excel (two worksheets)</li><li>Saved under <code>reports/&lt;timestamp&gt;/</code></li></ul></div>
  </div>
</section>

<section>
  <h2>Quick Start</h2>
  <pre><code>Unblock-File .\src\HyperV-ClusterHealth.ps1

# Option 1: prompt for guest creds (optional)
.\src\HyperV-ClusterHealth.ps1 -ClusterName "MyCluster"

# Option 2: pass creds explicitly (avoids prompts)
$cred = Get-Credential
.\src\HyperV-ClusterHealth.ps1 -ClusterName "MyCluster" -GuestCredential $cred</code></pre>
</section>

<section>
  <h2>Troubleshooting</h2>
  <ul>
    <li><strong>Guest OS shows “Unavailable”:</strong> Ensure guest service <em>Hyper-V Data Exchange Service</em> (<code>vmickvpexchange</code>) is Running + Automatic. With <code>-GuestCredential</code>, the script uses PowerShell Direct to start it and read OS details.</li>
    <li><strong>PowerShell 7 warning:</strong> Benign; Windows-only modules are imported via compatibility layer when needed.</li>
    <li><strong>Excel export:</strong> Requires the <code>ImportExcel</code> module; the script installs it for CurrentUser if missing.</li>
  </ul>
</section>

<section>
  <h2>Blog</h2>
  <p>Background article: <a href="https://www.cloudythoughts.cloud/?p=336&preview=true">Cloudy Thoughts</a></p>
</section>

<footer><p>MIT Licensed. &copy; 2025</p></footer>
</div>
</body>
</html>
