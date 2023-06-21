Sub AcceptAppointment()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myApptReq As Outlook.AppointmentItem
    
    Set myNameSpace = Outlook.Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myApptReq = myFolder.Items.Find("[MessageClass] = 'IPM.Appointment'")
    
    If Not myApptReq Is Nothing Then
        MsgBox "Appointment request found."
        If TypeName(myApptReq) = "AppointmentItem" Then
            MsgBox "Appointment request is a valid appointment item."
            myApptReq.Accept
            MsgBox "Appointment request accepted."
        Else
            MsgBox "Appointment request is not a valid appointment item."
        End If
    Else
        MsgBox "No appointment request found."
    End If
End Sub
