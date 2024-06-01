# Data Validation
## Code
```vb
Private Sub SetCellToFirstValidationRule(ByVal Cell As Range)
    If Cell.Count <> 1 Then Exit Sub
    If Cell.Validation.Type <> xlValidateList Then Exit Sub
    If Cell.Worksheet.ProtectContents = True And Cell.Locked = True Then Exit Sub
    
    Dim ValidationList As Variant
    ValidationList = Application.Evaluate(Cell.Validation.Formula1)
    
    Dim NewValue As Variant
    Select Case VarType(ValidationList)
        Case (vbArray + vbVariant)
            NewValue = ValidationList(1, 1)
        Case vbString
            NewValue = Split(ValidationList, ",")(0)
        Case Else
            Exit Sub
    End Select
    
    Cell.Value2 = NewValue
End Sub
```

## Notes
- Works with static/constant string Validation Lists, named ranges, same-worksheet ranges, and same-workbook/different-worksheet ranges.

## External Links
- https://learn.microsoft.com/en-us/office/vba/api/excel.validation
- https://learn.microsoft.com/en-us/office/vba/api/excel.xldvtype