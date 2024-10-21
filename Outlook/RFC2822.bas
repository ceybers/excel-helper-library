Option Explicit

Private Const MSG_CAPTION_FILTER As String = "Find Email by Message ID"
Private Const MSG_TEXT_FILTER As String = "Message ID of email to find:"

Private Const MSG_CAPTION_GETID As String = "Get Message ID of Email"
Private Const MSG_TEXT_GETID As String = "Message ID of selected email:"

Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

'@EntryPoint
Public Sub GetMessageIDofEmail()
    Dim MailItem As MailItem
    If Not TryGetSingleMailItem(MailItem) Then Exit Sub
    
    Dim MsgID As String
    MsgID = MailItem.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    InputBox MSG_TEXT_GETID, MSG_CAPTION_GETID, MsgID
End Sub

'@EntryPoint
Public Sub ShowEmailByMessageID()
    Dim MsgID As String
    MsgID = InputBox(MSG_TEXT_FILTER, MSG_CAPTION_FILTER)
    If MsgID = vbNullString Then Exit Sub

    Dim Query As String
    Query = GetDASLQuery(PR_INTERNET_MESSAGE_ID, "LIKE", MsgID)

    ' Search is restricted to the currently selected folder in the ActiveExplorer,
    ' and not the entire mailbox.
    Dim SearchResults As Items
    Set SearchResults = ActiveExplorer.CurrentFolder.Items.Restrict(Query)
    If SearchResults.Count = 1 Then SearchResults.Item(1).Display
End Sub

Private Function GetDASLQuery(ByVal Property As String, _
    ByVal Operator As String, ByVal Value As String) As String
    GetDASLQuery = "@SQL=" & _ 
        Chr$(34) & Property & Chr$(34) & _
        " " & Operator & " " & _
        "'" & Value & "'"
End Function

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