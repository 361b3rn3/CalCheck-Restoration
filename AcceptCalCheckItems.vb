Sub AcceptCalCheckItems()
    Dim CalCheckFolder As Outlook.Folder
    Dim CalCheckItem As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    
    ' Set the CalCheck folder where you want to accept the events
    Set objNamespace = Application.GetNamespace("MAPI")
    Set CalCheckFolder = objNamespace.GetDefaultFolder(olFolderInbox).Parent.Folders("CalCheck")
    
    ' Get the selected items in the CalCheck folder
    Set objExplorer = Application.ActiveExplorer
    Set objSelection = objExplorer.Selection
    
    ' Loop through each selected item in the CalCheck folder
    For Each CalCheckItem In objSelection
        ' Check if the item is an AppointmentItem or MeetingItem
        If TypeOf CalCheckItem Is Outlook.AppointmentItem Then
            ' Accept the event
            CalCheckItem.Respond (olResponseAccepted)
        ElseIf TypeOf CalCheckItem Is Outlook.MeetingItem Then
            ' Accept the meeting
            Dim MeetingItem As Outlook.MeetingItem
            Set MeetingItem = CalCheckItem
            MeetingItem.Respond (olMeetingAccepted)
            Set MeetingItem = Nothing
        End If
    Next CalCheckItem
    
    ' Clean up objects
    Set objNamespace = Nothing
    Set objExplorer = Nothing
    Set objSelection = Nothing
    Set CalCheckFolder = Nothing
    Set CalCheckItem = Nothing
End Sub
