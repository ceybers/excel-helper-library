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
## August 2025
- NewEnum issue
  - https://stackoverflow.com/questions/63848617/ (SO/Cristian Buse)
  - https://rubberduckvba.blog/2019/12/14/rubberduck-annotations/ (`@Enumerator`)
- `Dim foo as New foo`
  - tl;dr, will re-instantiate the object whenever it is referenced. Using `Set foo = Nothing` works up until another line tries to reference `foo` in which case it will automatically `Set foo = New foo`.
  - "An auto-instantiated object variable declaration at procedure scope changes how nulling the reference works, which can lead to unexpected behaviour."
  - https://stackoverflow.com/questions/8114684/
  - https://rubberduckvba.com/inspections/details/SelfAssignedDeclaration
- References: "Microsoft ActiveX Data Objects 6.1 Library".
  - Lets us use SQL in Excel, even if the target workbook is closed.
  - SELECT, UPDATE, INSERT work OK. However, DELETE does NOT work!
    - Pity because SQL JOINs are nicer than what we are currently using.
  - http://exceldevelopmentplatform.blogspot.com/2018/10/vba-microsoftaceoledb120-details.html
  - https://www.connectionstrings.com/ace-oledb-12-0/