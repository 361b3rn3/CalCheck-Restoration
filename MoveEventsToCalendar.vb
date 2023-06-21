Sub MoveEventsToCalendar()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim calCheckFolder As Outlook.MAPIFolder
    Dim calendarFolder As Outlook.MAPIFolder
    Dim eventItem As Outlook.AppointmentItem
    
    ' Create Outlook application and get the namespace
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get the "calcheck" folder
    Set calCheckFolder = olNamespace.GetDefaultFolder(olFolderInbox).Parent.Folders("CalCheck")
    
    ' Get the calendar folder
    Set calendarFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' Move each event item from "calcheck" to the calendar folder
    For Each eventItem In calCheckFolder.Items
        eventItem.Move calendarFolder
    Next eventItem
    
    ' Cleanup
    Set eventItem = Nothing
    Set calendarFolder = Nothing
    Set calCheckFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
    MsgBox "Event items moved to the calendar successfully!", vbInformation
End Sub
