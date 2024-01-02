Attribute VB_Name = "HashCellsUDF"
'@Folder("VBAProject")
Option Explicit

Private Const LEN_MD5_HASH As Long = 32
Private Const LEN_SHA256_HASH As Long = 64

'@Description "Calculates the MD5 hash of the values in a range of cells."
Public Function HashCellsMD5(ParamArray Range() As Variant) As Variant
    Dim ConcatenatedHashes As String
    Dim Cell As Range
    Dim Item As Variant
    
    For Each Item In Range
        If TypeOf Item Is Range Then
            For Each Cell In Item.Cells
                ConcatenatedHashes = ConcatenatedHashes & StringToMD5(CStr(Cell.Value2))
            Next Cell
        Else
            ConcatenatedHashes = ConcatenatedHashes & StringToMD5(CStr(Item))
        End If
    Next Item
    
    If Len(ConcatenatedHashes) > LEN_MD5_HASH Then
        ConcatenatedHashes = StringToMD5(ConcatenatedHashes)
    End If
    
    HashCellsMD5 = ConcatenatedHashes
End Function

'@Description "Calculates the SHA256 hash of the values in a range of cells."
Public Function HashCellsSHA256(ParamArray Range() As Variant) As Variant
    Dim ConcatenatedHashes As String
    Dim Cell As Range
    Dim Item As Variant
    
    For Each Item In Range
        If TypeOf Item Is Range Then
            For Each Cell In Item.Cells
                ConcatenatedHashes = ConcatenatedHashes & StringToSHA256(CStr(Cell.Value2))
            Next Cell
        Else
            ConcatenatedHashes = ConcatenatedHashes & StringToSHA256(CStr(Item))
        End If
    Next Item
    
    If Len(ConcatenatedHashes) > LEN_SHA256_HASH Then
        ConcatenatedHashes = StringToSHA256(ConcatenatedHashes)
    End If
    
    HashCellsSHA256 = ConcatenatedHashes
End Function

Private Function StringToMD5(ByVal Value As String) As String
    Static HashingObject As Object
    If HashingObject Is Nothing Then
        Set HashingObject = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    End If
    
    Dim BytesToEncode() As Byte
    BytesToEncode = StrConv(Value, vbFromUnicode)

    Dim EncodedBytes() As Byte
    EncodedBytes = HashingObject.ComputeHash_2((BytesToEncode))
    
    Dim Result As String
    Dim i As Long
    For i = 0 To UBound(EncodedBytes)
        Result = Result & Right$("0" & Hex$(AscB(MidB$(EncodedBytes, i + 1, 1))), 2)
    Next i
    
    StringToMD5 = Result
End Function

Private Function StringToSHA256(ByVal Value As String) As String
    Static HashingObject As Object
    If HashingObject Is Nothing Then
        Set HashingObject = CreateObject("System.Security.Cryptography.SHA256Managed")
    End If
    
    Dim BytesToEncode() As Byte
    BytesToEncode = StrConv(Value, vbFromUnicode)

    Dim EncodedBytes() As Byte
    EncodedBytes = HashingObject.ComputeHash_2((BytesToEncode))
    
    Dim Result As String
    Dim i As Long
    For i = 0 To UBound(EncodedBytes)
        Result = Result & Right$("0" & Hex$(AscB(MidB$(EncodedBytes, i + 1, 1))), 2)
    Next i
    
    StringToSHA256 = Result
End Function

Private Function BytesToBase64String(ByVal Bytes As Variant) As String
    Static Base64Object As Object
    If Base64Object Is Nothing Then
        Set Base64Object = CreateObject("MSXML2.DOMDocument")
        With Base64Object
            .LoadXML "<ROOT/>"
            .DocumentElement.DataType = "bin.base64"
        End With
    End If
    
    Base64Object.DocumentElement.NodeTypedValue = Bytes
    
    Dim Result As String
    Result = Replace(Base64Object.DocumentElement.Text, vbLf, vbNullString)
    
    BytesToBase64String = Result
End Function
