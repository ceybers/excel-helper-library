Attribute VB_Name = "HashCellsUDF"
'@IgnoreModule UseMeaningfulName
'@Folder "HashCells"
Option Explicit

'@Description "Calculates the SHA1 hash of a parameter array of Ranges."
'@EntryPoint
Public Function HashCellsSHA1(ParamArray Range() As Variant) As Variant
Attribute HashCellsSHA1.VB_Description = "Calculates the SHA1 hash of a parameter array of Ranges."
    Static Hasher As Object
    If Hasher Is Nothing Then
        Set Hasher = New SHA1Hasher
    End If
    
    HashCellsSHA1 = HashRangesGeneric(Hasher, Range)
End Function

'@Description "Calculates the SHA256 hash of a parameter array of Ranges."
'@EntryPoint
Public Function HashCellsSHA256(ParamArray Range() As Variant) As Variant
Attribute HashCellsSHA256.VB_Description = "Calculates the SHA256 hash of a parameter array of Ranges."
    Static Hasher As Object
    If Hasher Is Nothing Then
        Set Hasher = New SHA256Hasher
    End If
    
    HashCellsSHA256 = HashRangesGeneric(Hasher, Range)
End Function

'@Description "Calculates the MD5 hash of a parameter array of Ranges."
'@EntryPoint
Public Function HashCellsMD5(ParamArray Range() As Variant) As Variant
Attribute HashCellsMD5.VB_Description = "Calculates the MD5 hash of a parameter array of Ranges."
    Static Hasher As Object
    If Hasher Is Nothing Then
        Set Hasher = New MD5Hasher
    End If
    
    HashCellsMD5 = HashRangesGeneric(Hasher, Range)
End Function

Private Function HashRangesGeneric(ByVal Hasher As IHasher, ByVal Range As Variant) As Variant
    Dim Hashes() As String
    ReDim Hashes(0 To UBound(Range))
    
    Dim i As Long
    For i = 0 To UBound(Range)
        If IsEmpty(Range(i)) Then
            Hashes(i) = Hasher.ComputeHash(Chr$(0))
        Else
            Hashes(i) = HashRangeGeneric(Hasher, Range(i))
        End If
    Next i
    
    HashRangesGeneric = Hasher.ComputeHash(Join$(Hashes, vbNullString))
End Function

Private Function HashRangeGeneric(ByVal Hasher As IHasher, ByVal Range As Range) As String
    Dim vv As Variant
    vv = Range.Value2
    
    If Not IsArray(vv) Then
        HashRangeGeneric = Hasher.ComputeHash(CStr(vv))
        Exit Function
    End If
    
    Dim TotalHashes As Long
    TotalHashes = (((Range.Columns.Count * 2) + 1) * Range.Rows.Count) + 1
    
    Dim Hashes() As Variant
    ReDim Hashes(1 To TotalHashes)
    
    Dim CurrentHash As Long
    CurrentHash = 1
    
    Dim Row As Long
    Dim Column As Long
    For Row = 1 To Range.Rows.Count
        For Column = 1 To Range.Columns.Count
            Hashes(CurrentHash) = Hasher.ComputeHash(CStr(vv(Row, Column)))
            Hashes(CurrentHash + 1) = Chr$(31)
            CurrentHash = CurrentHash + 2
        Next Column
        
        Hashes(CurrentHash) = Chr$(30)
        CurrentHash = CurrentHash + 1
    Next Row
    Hashes(CurrentHash) = Chr$(29)
    
    HashRangeGeneric = Hasher.ComputeHash(Join(Hashes, vbNullString))
End Function
