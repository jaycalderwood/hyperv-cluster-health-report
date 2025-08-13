# Troubleshooting

## Guest OS shows "Unavailable"
- Ensure the guest Windows service **Hyper-V Data Exchange Service** (`vmickvpexchange`) is Running and Automatic.
- Run with `-GuestCredential` so the script can use PowerShell Direct to start the service and read OS details.

## PowerShell 7 deserialized objects
- Benign for `FailoverClusters` in PS7; the script handles module import via compatibility.

## ImportExcel issues
- Make sure PSGallery is reachable; the script tries `Install-Module ImportExcel -Scope CurrentUser` if missing.
