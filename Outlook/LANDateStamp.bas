Option Explicit

Private Const MSG_CAPTION_LAN As String = "Get Mail LAN ID & Date"
Private Const MSG_TEXT_LAN As String = "LAN ID & Date stamp:"

Private Const DATE_TIME_FORMAT As String = "ddd yyyy/mm/dd hh:mm"

'@EntryPoint
Public Sub ShowLANDateStamp()
    Dim MailItem As MailItem
    If Not TryGetSingleMailItem(MailItem) Then Exit Sub

    ' Sender will be Nothing if e.g. new draft email.
    If MailItem.Sender Is Nothing Then Exit Sub

    Dim Result As String
    If CurItem.Sender.GetExchangeUser Is Nothing Then
        Result = CurItem.SenderEmailAddress & " " & GetTimeOfMailItem(CurItem)
    Else
        Result = CurItem.Sender.GetExchangeUser.Alias & " " & GetTimeOfMailItem(CurItem)
    End If

    InputBox MSG_CAPTION_LAN, MSG_CAPTION, Result
End Sub

Private Function TryGetSingleMailItem(ByRef OutMailItem As MailItem) As Boolean
    Dim SomeItem As Object

    If TypeOf ActiveWindow Is Inspector Then
        Set SomeItem = ActiveInspector.CurrentItem
    ElseIf TypeOf ActiveWindow Is Explorer Then
        If ActiveWindow.Selection.Count >= 1 Then
            Set SomeItem = ActiveWindow.Selection.Item(1)
        End If
    End If

    If TypeOf SomeItem Is MailItem Then
        Set OutMailItem = SomeItem
        TryGetSingleMailItem = True
    End If
End Function

Private Function GetTimeOfMailItem(ByVal MailItem As MailItem) As String
    GetTimeOfMailItem = Format$(CurItem.ReceivedTime, DATE_TIME_FORMAT)
End Function