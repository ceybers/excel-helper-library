# Developer Notes for Excel VBA
## Notes
### String Concatenation and Performance
- Use Arrays and `Join()` instead of concatenating strings (e.g., `foobar = foobar & xyz`) when having to join large groups of data.
### UDFs and Static Objects
- Use the `Static foobar as Object: If foobar Is Nothing: Set foobar = CreateFoobar()` pattern to avoid having to instantiate Objects for each cell that calls a UDF.
## API References
## External References
- [Keywords - Visual Basic | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/keywords/)
## Additional Controls
### Status Bar
- `Microsoft StatusBar Control, version 6.0`
- `C:\Windows\system32\MSCOMCTL.OCX`
- `Me.StatusBar1.Panels.Add Text:="hello"`
## Toolbar
- MSComctlLib.ButtonStyleConstants
- Images seem to not be working in 64-bit Excel.
## Schema Change Tracking
- [Dataedo](https://dataedo.com/asset/img/docs/8_0/schema_change_tracking_overview.png)