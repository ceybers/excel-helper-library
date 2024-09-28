# `TryGetUsableSelection()`

## Description
Returns `TRUE` and sets the `Range` Out parameter if the `Selection` object passes the criteria below.

## Criteria
- The worksheet must be unprotected.
- A selection of cells spanning two or more rows must be selected.
- Neither entire column(s) nor entire row(s) may be selected.
- Multiple selections (`Ctrl`+select, non-contiguous selections) are allowed as long as one or more of the selections span more than one row.
  - The selections that span only one row will be ignored.
  - If any of the selections fail the check for entire column/rows, they will all be failed.

## Behaviour
- If a `Selection` fails any of the criteria above, a `MsgBox` will be displayed to the user.

## Function Signature
```vb
Public Function TryGetUsableSelection(ByVal Caption As String, ByRef OutRange As Range) As Boolean
```