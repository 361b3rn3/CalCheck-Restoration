# Load the Outlook COM object
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Create an instance of the Outlook Application
$outlook = New-Object -ComObject Outlook.Application

# Get the default Outlook namespace
$namespace = $outlook.GetNamespace("MAPI")

# Specify the name of the folder containing the appointments (excluding Calendar)
$folderName = "Your Folder Name"

# Get the specified folder
$folder = $namespace.Folders | Where-Object { $_.Name -eq $folderName }

# Check if the folder exists
if ($folder -eq $null) {
    Write-Host "Folder '$folderName' not found."
    exit
}

# Get the items (appointments) from the folder
$appointments = $folder.Items

# Loop through each appointment and perform desired actions
foreach ($appointment in $appointments) {
    # Access appointment properties as needed
    $subject = $appointment.Subject
    $start = $appointment.Start
    $end = $appointment.End

    # Do something with the appointment (e.g., display information)
    Write-Host "Subject: $subject"
    Write-Host "Start: $start"
    Write-Host "End: $end"

    # Accept the appointment
    $appointment.Respond(3)  # 3 corresponds to the 'olResponseAccepted' constant

    # Save changes to the appointment
    $appointment.Save()
}

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($appointments) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($folder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
