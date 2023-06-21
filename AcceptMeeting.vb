Sub AcceptMeeting()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myMtgReq As Outlook.MeetingItem
    Dim myAppt As Outlook.AppointmentItem
    Dim myMtg As Outlook.MeetingItem
    
    Set myNameSpace = Outlook.Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myMtgReq = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'")
    
    If Not myMtgReq Is Nothing Then
        MsgBox "Meeting request found."
        If TypeName(myMtgReq) = "MeetingItem" Then
            MsgBox "Meeting request is a valid meeting item."
            Set myAppt = myMtgReq.GetAssociatedAppointment(True)
            Set myMtg = myAppt.Respond(olResponseAccepted, True)
            myMtg.Send
            MsgBox "Meeting request accepted and sent."
        Else
            MsgBox "Meeting request is not a valid meeting item."
        End If
    Else
        MsgBox "No meeting request found."
    End If
End Sub
