Sub AcceptCalCheckItems()
    Dim CalCheckFolder As Outlook.Folder
    Dim CalCheckItem As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    
    ' Set the CalCheck folder where you want to accept the events
    Set objNamespace = Application.GetNamespace("MAPI")
    On Error Resume Next
    Set CalCheckFolder = objNamespace.GetDefaultFolder(olFolderInbox).Parent.Folders("CalCheck")
    On Error GoTo 0
    
    ' Check if the CalCheck folder exists
    If CalCheckFolder Is Nothing Then
        MsgBox "CalCheck folder does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Get the selected items in the CalCheck folder
    Set objExplorer = Application.ActiveExplorer
    Set objSelection = objExplorer.Selection
    
    ' Loop through each selected item in the CalCheck folder
    For Each CalCheckItem In objSelection
        ' Display a message box indicating the type of item that is going to be processed
        Dim ItemType As String
        If TypeOf CalCheckItem Is Outlook.AppointmentItem Then
            ItemType = "AppointmentItem"
        ElseIf TypeOf CalCheckItem Is Outlook.MeetingItem Then
            ItemType = "MeetingItem"
        ElseIf TypeOf CalCheckItem Is Outlook.MailItem Then
            ItemType = "MailItem"
        Else
            ItemType = "Unsupported item type: " & TypeName(CalCheckItem)
        End If
        MsgBox "Processing item: " & ItemType, vbInformation
        
        ' Process the item based on its type
        If TypeOf CalCheckItem Is Outlook.AppointmentItem Then
            ' Accept the event
            Dim AppointmentItem As Outlook.AppointmentItem
            Set AppointmentItem = CalCheckItem
            AppointmentItem.Respond (olResponseAccepted)
            Set AppointmentItem = Nothing
        ElseIf TypeOf CalCheckItem Is Outlook.MeetingItem Then
            ' Accept the meeting
            Dim MeetingItem As Outlook.MeetingItem
            Set MeetingItem = CalCheckItem
            MeetingItem.Respond (olMeetingAccepted)
            Set MeetingItem = Nothing
        ElseIf TypeOf CalCheckItem Is Outlook.MailItem Then
            ' Accept regular mail items
            Dim MailItem As Outlook.MailItem
            Set MailItem = CalCheckItem
            MailItem.ReplyAll
            Set MailItem = Nothing
        End If
    Next CalCheckItem
    
    ' Clean up objects
    Set objNamespace = Nothing
    Set objExplorer = Nothing
    Set objSelection = Nothing
    Set CalCheckFolder = Nothing
    Set CalCheckItem = Nothing
End Sub
