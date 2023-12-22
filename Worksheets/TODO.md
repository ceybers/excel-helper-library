
# Worksheet Helpers TODO
## Implement
- `TryGetWorksheet(ByVal Parent as Object, ByVal WorksheetName as String, ByRef OutWorksheet as Worksheeet) as Boolean`
    - Use "like" keyword instead of `=` equals operator.
    Parent object can be one of multiple types.
    - ...in `Collection<IWorksheeet>`
        - User is responsible for ensuring all items in Collection can be cast to class of Workbook.
    - ...in `Worksheets` property of a `Workbook`
    - ...in `Workbook` object
    - ...in `Application` object (i.e., all the open workbooks)
        - Name is guaranteed to be unique within one workbook.
        - Name is not guaranteed to be unique in multiple workbooks - do we fail if more than one result? Only return true if exactly one result? Return first result? Or change signature and return a collection?
- `GetDictionaryOfWorksheets(ByVal Parent as Object)`
    -   Alt, `GetWorksheetsAsDictionary()`
    - Same Strategy Pattern as above.
    - Key value *cannot* be `Worksheet.Name`, it has to be the fully qualified name.
## Thoughts
- FQN for Worksheets?
    `Workbook.Name:Worksheet.Name`
    `Workbook.Name:Worksheet.Name:ListObject.Name`
    Use `.Name` and not full name with the path. If we do, we must remember to handle the colon in `C:\` or `https://`. Network shares won't have one, so we cannot assuume one (`\\server\folder\file.xlsx`)
- Don't need a non `TryGet...` pattern because it would be the same as the normal iterator/item accessor in a collection.
    - e.g., `Worksheets.Item(worksheet_name)`.
## Other pairs of classes
### ListObjects
- ListObject in Range (Selection is also a Range)
- ListObject in Worksheets(.ListObjects)
- ListObject in Workbook(.Worksheets)
- Collections of ListObject
- Collections of ListObjects properties (flatten)
- Collections of Worksheets, Workbooks
- ListObjects in Applications
- Interesting function would be `GetExactlyOneListObject()` that only returns True if a Worksheet has one ListObject.
### ListColumn in ListObject
- As above
- Also:
    - ListColumn from Range (must return exactly one)
    - ListColumns from Range (must handle non-contiguous/multiple areas, but in same ListObject)
    - Get Index from List Column Name (ListColumn object has an Index property already)
        - Useful because Filters can work on field numbers.
    - Get List Column by Index Number.
- Also also:
    - ListRows from Range
    - If the area is non-contiguous we could return one collection of ListRows and one collection of ListColumns.
        - Would most likely require a custom return object type.
### Workbooks in Application