$cred = Get-Credential
.\src\HyperV-ClusterHealth.ps1 -ClusterName "MyCluster" -GuestCredential $cred
