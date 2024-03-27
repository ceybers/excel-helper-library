VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA1Hasher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "HashCells"
Option Explicit
Implements IHasher

Private mHashingObject As Object
Private Const ENCODED_BYTES_LEN As Long = 19

Private Sub Class_Initialize()
    Set mHashingObject = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
End Sub

Private Function IHasher_ComputeHash(ByVal Value As String) As String
    Dim BytesToEncode() As Byte
    BytesToEncode = StrConv(Value, vbFromUnicode)

    Dim EncodedBytes() As Byte
    EncodedBytes = mHashingObject.ComputeHash_2((BytesToEncode))
    
    Dim Result(0 To ENCODED_BYTES_LEN) As String
    Dim i As Long
    For i = 0 To ENCODED_BYTES_LEN
        Result(i) = Right$("0" & Hex$(AscB(MidB$(EncodedBytes, i + 1, 1))), 2)
    Next i
   
    IHasher_ComputeHash = Join(Result, vbNullString)
End Function
