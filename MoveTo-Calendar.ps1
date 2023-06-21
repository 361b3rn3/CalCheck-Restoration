Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Create an instance of Outlook
$outlook = New-Object -ComObject Outlook.Application

# Get the default namespace and the "calcheck" folder
$namespace = $outlook.GetNamespace("MAPI")
$calCheckFolder = $namespace.GetDefaultFolder(6).Folders["calcheck"]

# Get the calendar folder
$calendarFolder = $namespace.GetDefaultFolder(9)

# Move each event item from "calcheck" to the calendar folder
$calCheckFolder.Items | ForEach-Object {
    $_.Move($calendarFolder)
}

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($calCheckFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendarFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Event items moved to the calendar successfully!"
