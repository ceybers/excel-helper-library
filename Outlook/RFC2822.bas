Public Sub ShowLANDateStamp()
    Dim CurItem As Object
    If TypeOf ActiveWindow Is Inspector Then
        Set CurItem = ActiveInspector.CurrentItem
    Else
        Set CurItem = ActiveWindow.Selection.Item(1)
    End If

    If CurItem.Sender Is Nothing Then Exit Sub

    Dim Result As String
    If CurItem.Sender.GetExchangeUser Is Nothing Then
        Result = CurItem.SenderEmailAddress & " " & Format(CurItem.ReceivedTime, "ddd yyyy/mm/dd hh:mm")
    Else
        Result = CurItem.Sender.GetExchangeUser.Alias & " " & Format(CurItem.ReceivedTime, "ddd yyyy/mm/dd hh:mm")
    End If

    InputBox "LAN ID & Date", "CAE Macro", Result

    Set CurItem = Nothing
End Sub

Public Sub ShowMessageID()
    Dim CurItem As Object
    If TypeOf ActiveWindow Is Inspector Then
        Set CurItem = ActiveInspector.CurrentItem
    Else
        Set CurItem = ActiveWindow.Selection.Item(1)
    End If
    
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = CurItem.PropertyAccessor

    Dim GetInetHeaders As String
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)

    Dim MsgID As String
    MsgID = CurItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
    InputBox "MessageID", "CAE Macro", MsgID

    Set CurItem = Nothing
End Sub

Public Sub FilterMessageID()
    Dim MsgID As String
    MsgID = InputBox("MsgID to filter by:", "Filter by Message ID")

    Dim Query As String
    Query = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x1035001F"" LIKE '" & MsgID & "'"

    Dim Items As Items
    Set Items = ActiveExplorer.CurrentFolder.Items.Restrict(Query)
    If Items.Count = 1 Then Items.Item(1).Display
End Sub