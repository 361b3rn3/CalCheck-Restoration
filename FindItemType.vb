Sub FindItemType()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    
    Set myNameSpace = Outlook.Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) ' Change the folder as needed
    Set myItems = myFolder.Items
    
    ' Loop through each item in the folder
    For Each myItem In myItems
        ' Determine the item type and display a message box
        Select Case TypeName(myItem)
            Case "MailItem"
                MsgBox "MailItem: " & myItem.Subject
            Case "AppointmentItem"
                MsgBox "AppointmentItem: " & myItem.Subject
            Case "MeetingItem"
                MsgBox "MeetingItem: " & myItem.Subject
            Case "TaskItem"
                MsgBox "TaskItem: " & myItem.Subject
            Case Else
                MsgBox "Other Item Type: " & TypeName(myItem)
        End Select
    Next myItem
End Sub
