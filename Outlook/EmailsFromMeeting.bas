Public Sub GetEmailAddressFromMeeting()
    If Application.ActiveInspector Is Nothing Then Exit Sub
    If Not TypeOf Application.ActiveInspector.CurrentItem Is AppointmentItem Then Exit Sub

    Dim AppointmentItem As AppointmentItem
    Set AppointmentItem = Application.ActiveInspector.CurrentItem

    Debug.Print "Subject: "; AppointmentItem.Subject
    Debug.Print "Time: "; AppointmentItem.Start; " to "; AppointmentItem.End
    Debug.Print vbNullString

    Dim Recipient As Recipient
    For Each Recipient In AppointmentItem.Recipients
        Debug.Print Recipient.Name; " <"; GetSmtpAddress(Recipient); ">"
    Next Recipient
End Sub

Private Function GetSmtpAddress(ByVal Recipient As Recipient) As String
    Select Case Recipient.AddressEntry.AddressEntryUserType
        Case olSmtpAddressEntry
            GetSmtpAddress = Recipient.AddressEntry.Address
        Case olExchangeUserAddressEntry
            GetSmtpAddress = Recipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        Case olExchangeDistributionListAddressEntry
            GetSmtpAddress = Recipient.AddressEntry.GetExchangeDistributionList.PrimarySmtpAddress
    End Select
End Function