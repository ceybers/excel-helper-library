Option Explicit

'@EntryPoint
Public Sub GetEmailAddressFromMeeting()
    If Application.ActiveInspector Is Nothing Then Exit Sub
    If Not TypeOf Application.ActiveInspector.CurrentItem Is AppointmentItem Then Exit Sub

    Dim AppointmentItem As AppointmentItem
    Set AppointmentItem = Application.ActiveInspector.CurrentItem

    Dim RecipientCount As Long
    RecipientCount = AppointmentItem.Recipients.Count

    Dim Output() as String
    ReDim Output(1 to 3 + RecipientCount)

    Output(1) = "Subject: " & AppointmentItem.Subject
    Output(2) = "Time: " & AppointmentItem.Start & " to " & AppointmentItem.End
    Output(3) = "Recipients:" & RecipientCount

    Dim i as Long
    For i = 1 To RecipientCount
        Output(3 + i) = Recipient.Name & " <" & GetSmtpAddress(Recipient) & ">"
    Next i

    Debug.Print Join$(Output, vbCrLf)
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