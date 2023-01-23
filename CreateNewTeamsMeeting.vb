Sub CreateNewTeamsMeeting()
    Dim startDate As Date
    Dim endDate As Date
    Dim meetingBody As String

    ' Get today's date
    startDate = Date

    ' Set end date as 7 days from today
    endDate = DateAdd("d", 7, startDate)

    ' Create new Teams meeting
    Dim teamsMeeting As Outlook.AppointmentItem
    Set teamsMeeting = Application.CreateItem(olAppointmentItem)

    ' Set meeting properties
    teamsMeeting.Start = startDate
    teamsMeeting.End = endDate
    teamsMeeting.Subject = "New Teams Meeting"
    teamsMeeting.Location = "Microsoft Teams"
    teamsMeeting.MeetingStatus = olMeeting
    teamsMeeting.Body = "This is a new Teams meeting"

    ' Display the meeting for editing
    teamsMeeting.Display

    ' Copy the body of the meeting to the clipboard
    meetingBody = teamsMeeting.Body
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText meetingBody
    MSForms_DataObject.PutInClipboard

    ' Save the meeting after editing if needed
    teamsMeeting.Save
End Sub
