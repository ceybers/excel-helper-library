Option Explicit

Private Const EXPORT_FOLDER = "C:\Users\User\Desktop\New Folder"
Private Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
Private Const PR_MIME_TYPE As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"

Public Sub ExportAttachmentsToFolders()
    DoSaveAttachmentsToFolders
End Sub

Private Sub DoSaveAttachmentsToFolders()
    CreateFolder EXPORT_FOLDER

    Dim EmailsProcessed As Long
    Dim AttachmentsSaved As Long

    Dim SelectedObject As Object
    For Each SelectedObject In ActiveWindow.Selection
        If TypeOf SelectedObject Is MailItem Then
            AttachmentsSaved = AttachmentsSaved + ProcessMailItem(SelectedObject)
            EmailsProcessed = EmailsProcessed + 1
        End If
    Next SelectedObject

    Dim MsgboxMessage As String
    MsgboxMessage = "Exported " & AttachmentsSaved & " attachments from " & EmailsProcessed & " email(s)."
    MsgBox MsgboxMessage, vbInformation + vbOKOnly, "Save Attachmetns to Folders"

    Shell "C:\WINDOWS\explorer.exe """ & EXPORT_FOLDER & "", vbNormalFocus
End Sub

Private Function ProcessMailItem(ByVal MailItem As MailItem) As Long
    Dim AttachmentsToSave As Collection
    Set AttachmentsToSave = New Collection

    Dim Attachment As Attachment
    For Each Attachment In MailItem.AttachmentsSaved
        Dim MimeType As String
        MimeType = Attachment.PropertyAccessor.GetProperty(PR_MIME_TYPE)

        If Left$(MimeType, 5) <> "image" Then
            AttachmentsToSave.Add Attachment
        End If
    Next Attachment

    If AttachmentsToSave.Count = 0 Then Exit Function

    Dim DestinationFolder As String
    DestinationFolder = GetFolderName(MailItem)
    CreateFolder DestinationFolder

    For Each Attachment In AttachmentsToSave
        Attachment.SaveAsFile (DestinationFolder & "\" & Attachment.DisplayName)
    Next Attachment

    ProcessMailItem = AttachmentsToSave.Count
End Function

Private Function GetFolderName(ByVal MailItem As MailItem) As String
    Dim exUsr As ExchangeUser
    Dim formattedUsr As String
    Set exUsr = MailItem.Sender.GetExchangeUser
    If exUsr Is Nothing Then
        formattedUsr = EncodeURL(MailItem.SenderEmailAddress)
    Else
        formattedUsr = exUsr.Alias
    End If

    GetFolderName = EXPORT_FOLDER & Left$(EncodeURL(MailItem.Subject), 64) & " " & _
        formattedUsr & " " & EncodeURL(MailItem.ReceivedTime)
End Function

Private Function CreateFolder(ByVal folderName As String) As Integer
    Dim fsoObject As Object
    Set fsoObject = CreateObject("Scripting.FileSystemObject")
    If fsoObject.FolderExists(folderName) Then
        CreateFolder = 2
        Exit Function
    Else
        Call fsoObject.CreateFolder(folderName)
        CreateFolder = 1
        Exit Function
    End If
    CreateFolder = 0
End Function

Private Function EncodeURL(ByVal url As String) As String
    ' Based on https://stackoverflow.com/a/49502477
    Dim output As String

    Dim i As Integer
    For i = 1 To Len(url)
        Dim cur As Long
        cur = Asc(Mid(url, i, 1))
        Select Case cur
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95
                output = output + Mid(url, i, 1)
            Case Else
                output = output + "_"
        End Select
    Next i

    EncodeURL = output
End Function

