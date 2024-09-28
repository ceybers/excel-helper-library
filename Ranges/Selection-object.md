# `Selection` object in Excel
- If one or more cells are selected, the type of the `Selection` object is `Object/Range`.
  - PivotTables and ListObjects (Tables) return a `Object/Range`.
  - That `Range` object has a `.ListObject` and a `'.PivotTable` property.
- If one chart is selected, it returns a `Object/ChartArea`. Multiple charts returns `Object/DrawingObjects`.
- If one picture is selected, it returns a `Object/Picture`. Multiple pcitures return a `Object/DrawingObjects`.
- If one shape is selected, it is, e.g., `Object/Rectangle`.
  - `TypeOf Selection is Rectangle` = TRUE
  - `TypeOf Selection is Shape` = FALSE
  - `TypeOf Selection is ShapeRange` = FALSE
  - Has a ShapeRange property
  - `TypeOf Selection.Parent is Worksheet`
- Multiple shapes returns a `Object/DrawingObjects`.
- Under what situation does it return `Nothing`?
  - Even when a worksheet is protected and `Select cells` is disabled, `Selection` references the cell in the Name Box (to the left of the formula bar). 

# External References
- [Selection Object | Microsoft Learn](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa223084(v=office.11))
- [Application.Selection property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.application.selection)