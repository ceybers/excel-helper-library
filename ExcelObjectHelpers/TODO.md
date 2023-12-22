
# Excel Object Helpers TODO
# Worksheets
## TryGetWorksheet
`TryGetWorksheet(ByVal Parent as Object, ByVal WorksheetName as String, ByRef OutWorksheet as Worksheeet) as Boolean`
- [x] Use "like" keyword instead of `=` equals operator.
- Parent object can be one of multiple types.
  - [x] ...in `Collection<IWorksheeet>`
      - User is responsible for ensuring all items in Collection can be cast to class of Workbook.
  - [x] ...in `Worksheets` property of a `Workbook`
  - [x] ...in `Workbook` object
      - Name is guaranteed to be unique within one workbook.
  - [x] ...in `Application` object (i.e., all the open workbooks)
      - Name is not guaranteed to be unique in multiple workbooks - do we fail if more than one result? Only return true if exactly one result? Return first result? Or change signature and return a collection?
## GetDictionaryOfWorksheets
`GetDictionaryOfWorksheets(ByVal Parent as Object)`
-   Alt, `GetWorksheetsAsDictionary()`
- ~~Same Strategy Pattern as above.~~
- Key value *cannot* be `Worksheet.Name`, it has to be the fully qualified name.
## Thoughts
- [x] FQN for Worksheets?
    `Workbook.Name:Worksheet.Name`
    `Workbook.Name:Worksheet.Name:ListObject.Name`
    Use `.Name` and not full name with the path. If we do, we must remember to handle the colon in `C:\` or `https://`. Network shares won't have one, so we cannot assuume one (`\\server\folder\file.xlsx`)
- [x] Don't need a non `TryGet...` pattern because it would be the same as the normal iterator/item accessor in a collection.
    - e.g., `Worksheets.Item(worksheet_name)`.
# ListObjects
- [ ] ListObject in Range (Selection is also a Range)
- [x] ListObject in Worksheets(.ListObjects)
- [x] ListObject in Workbook(.Worksheets)
- [x] Collections of ListObject
- [ ] Collections of ListObjects properties (flatten)
- [x] Collections of Worksheets, Workbooks
- [x] ListObjects in Applications
- Interesting function would be `GetExactlyOneListObject()` that only returns True if a Worksheet has one ListObject.
# ListColumns
- [ ] ListColumn from Range (must return exactly one)
- [ ] ListColumns from Range (must handle non-contiguous/multiple areas, but in same ListObject)
- [ ] Get Index from List Column Name (ListColumn object has an Index property already)
    - Useful because Filters can work on field numbers.
- [ ] Get List Column by Index Number.
# ListRows
- [ ] ListRows from Range
- [ ] If the area is non-contiguous we could return one collection of ListRows and one collection of ListColumns.
    - Would most likely require a custom return object type.
# See Also
- [Array class helpers](../Arrays/ArrayHelpers.md)
- [Collection class helpers](../Collections/CollectionHelpers.md)
- [ListColumn class helpers](../ListColumns/ListColumnHelpers.md)
- [ListObject class helpers](../ListObjects/ListObjectHelpers.md)
- [Workbook class helpers](../Workbooks/WorkbookHelpers.md)
- [Worksheet class helpers](../Worksheets/WorksheetHelpers.md)