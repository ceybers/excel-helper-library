# `Selection` object in Excel
- If one or more cells are selected, the type of the `Selection` object is `Object/Range`.
  - PivotTables and ListObjects (Tables) return a `Object/Range`.
  - That `Range` object has a `.ListObject` and a `'.PivotTable` property.
  - Non-contiguous ranges will return a `Range` whose `Area` property has multiple items.
    - The `.Value2` property of the top `Range` will only return the `.Value2` for the first item.
    - The `.Cells()` property will loop through every cell in all areas. Each item will be a `Range` object of `1×1` size.
  -  If the size of a `Selection` is unknown, make sure to handle the case of it being either a `1×1` cell or a 2-dimensional array, either `1×n`, `m×1` or `m×n` size.
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