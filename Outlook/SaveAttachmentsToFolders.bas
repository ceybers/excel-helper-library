Option Explicit

Private Const FOLDER_PREFIX As String = "OutlookAttachments_"
Private Const MSG_CAPTION As String = "Save attachments in selected emails to folder"
Private Const ERR_MSG_CANTCREATEFOLDER As String = "Could not create a folder to save the attachments in."
Private Const PR_MIME_TYPE As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"

Private Type ExporterResult
    AttachmentsSaved As Long
    EmailsProcessed As Long
End Type

'@EntryPoint
Public Sub ExportAttachmentsToFolders()
    Dim ExportFolder As String
    ExportFolder = GetTemporaryFolderPath()

    If Not TryCreateFolder(ExportFolder) Then
        MsgBox ERR_MSG_CANTCREATEFOLDER, vbExclamation + vbOKOnly, MSG_CAPTION
    End If

    Dim Result As ExporterResult
    Result = ProcessSelection(ExportFolder)
    DisplayResult Result

    OpenFolderInExplorer ExportFolder
End Sub

Private Function ProcessSelection(ByVal Path As String) As ExporterResult
    Dim ThisItem As Object
    For Each ThisItem In ActiveWindow.Selection
        Dim Result As ExporterResult
        Result = ProcessMailItem(ThisItem, Path)

        With ProcessSelection
            .AttachmentsSaved = .AttachmentsSaved + Result.AttachmentsSaved
            .EmailsProcessed = .EmailsProcessed + Result.EmailsProcessed
        End With
    Next ThisItem
End Function

Private Function ProcessMailItem(ByVal Object As Object, ByVal Path As String) As ExporterResult
    If Not TypeOf Object Is MailItem Then Exit Function
    ProcessMailItem.EmailsProcessed = 1

    Dim AttachmentsToSave As Collection
    Set AttachmentsToSave = GetAttachmentsToSave(Object)

    ' Don't create a folder for the email if there are no attachments to save.
    If AttachmentsToSave.Count = 0 Then Exit Function

    Dim PathForEmail As String
    PathForEmail = Path & "\" & GetSubfolderNameForEmail(Object)
    If Not TryCreateFolder(PathForEmail) Then Exit Function

    With ProcessMailItem
        .AttachmentsSaved = TrySaveAttachments(AttachmentsToSave, PathForEmail)
    End With
End Function

' Returns a collection containing all the attachments that pass the condition test.
Private Function GetAttachmentsToSave(ByVal MailItem As MailItem) As Collection
    Set GetAttachmentsToSave = New Collection

    Dim Attachment As Attachment
    For Each Attachment In MailItem.AttachmentsSaved
        If TestIfAttachmentMustBeSaved(Attachment) Then
            GetAttachmentsToSave.Add Attachment
        End If
    Next Attachment
End Function

Private Function TestIfAttachmentMustBeSaved(ByVal Attachment As Attachment) As Boolean
    Dim MimeType As String
    MimeType = Attachment.PropertyAccessor.GetProperty(PR_MIME_TYPE)

    If Left$(MimeType, 5) <> "image" Then
        TestIfAttachmentMustBeSaved = True
    End If
End Function

' Saves all the attachments in the given collection to the give folder.
' Returns a long with the number of attachments saved.
Private Function TrySaveAttachments(ByVal AttachmentsToSave As Collection, _
    ByVal PathForEmail As String) As Long

    Dim Attachment As Attachment
    For Each Attachment In AttachmentsToSave
        Dim AttachmentPath As String
        AttachmentPath = PathForEmail & "\" & Attachment.DisplayName

        Attachment.SaveAsFile AttachmentPath

        TrySaveAttachments = TrySaveAttachments + 1
    Next Attachment
End Function

Private Sub DisplayResult(ByRef Result As ExporterResult)
    Dim MsgboxMessage As String
    MsgboxMessage = "Exported " & _
        Result.AttachmentsSaved & " attachments from " & _
        Result.EmailsProcessed & " email(s)."

    MsgBox MsgboxMessage, vbInformation + vbOKOnly, MSG_CAPTION
End Sub

Private Sub OpenFolderInExplorer(ByVal Path As String)
    Shell "C:\WINDOWS\explorer.exe """ & Path, vbNormalFocus
End Sub

' Returns a folder name for the attachments contained in a specific email.
' Format is `Subject{36} Sender DateTime`, with characters sanitised
' to prevent invalid folder names.
Private Function GetSubfolderNameForEmail(ByVal MailItem As MailItem) As String
    Dim ExchangeUser As ExchangeUser
    Set ExchangeUser = MailItem.Sender.GetExchangeUser

    Dim FormattedUsername As String
    If ExchangeUser Is Nothing Then
        FormattedUsername = SanitiseString(MailItem.SenderEmailAddress)
    Else
        FormattedUsername = ExchangeUser.Alias
    End If

    GetSubfolderNameForEmail = _
        Left$(SanitiseString(MailItem.Subject), 64) & " " & _
        FormattedUsername & " " & _
        GetTimeISO8601(MailItem.ReceivedTime)
End Function

' Returns True if the folder was succesfully created or if it already existed.
' Returns False if folder does not exist and could not be created.
Private Function TryCreateFolder(ByVal Path As String) As Boolean
    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

    If Not FileSystemObject.FolderExists(Path) Then
        FileSystemObject.CreateFolder Path
    End If

    TryCreateFolder = FileSystemObject.FolderExists(Path)
End Function

' Returns the input string filtered to only contain the below characters:
' [A-z0-9\-\.\@]
' All other characters are replaced by an underscore.
Private Function SanitiseString(ByVal Value As String) As String
    Dim Bytes() As Byte
    Bytes = Value
    
    Dim i As Long
    For i = 0 To UBound(Bytes) Step 2
        Select Case Bytes(i)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 64
                ' Do Nothing
            Case Else
                Bytes(i) = 95
        End Select
    Next i
    
    SanitiseString = Bytes
End Function

Private Function GetTimeISO8601(ByVal DateTime As Date) As String
    If Not IsDate(DateTime) Then Exit Function
    GetTimeISO8601 = Format$(DateTime, "yyyymmddThhMMss")
End Function

' Returns a path where folders will be created for each selected email
' that contains attachments that are succesfully saved.
Private Function GetTemporaryFolderPath() As String
    GetTemporaryFolderPath = Environ$("TMP") & "\" & FOLDER_PREFIX & GetTimeISO8601(Now())
End Function