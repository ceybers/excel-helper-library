# SettingsModel
## Description
- A model to store persistent data in a Workbook, in the form of: settings, flags, and collections.
  - Settings are key value pairs, with the values always being strings. Default value is `vbNullString`.
  - Flags are boolean keys that can be either `True` or `False`. Default value is `False`.
  - Collections are collections of Strings. Default is an empty `Collection` object.
- All three of these can be stored either at the workbook-level (singleton), or at the table-level (supporting multiple tables).
- The tables are not linked or limited to the actual ListObjects in the workbook.
- If a key does not exist, the getter returns the default value. The setter will automatically insert the key if it doesn't exist, and will update it if it does (i.e., Upsert).
- If no settings model already exists, using the Create method will create an empty one.

## TODO
- [x] Basic user-level persistence in My Documents folder
  - [x] Flag (boolean) support
  - [ ] Setting (string) support
  - [ ] Collection support
  - [ ] Most Recently Used (MRU) support
  - [ ] Manual Save mode
- [x] Workbook-level and ListObject-level persistence via CustomXMLPart object
  - [x] Flag (boolean) support
  - [x] Setting (string) support
  - [x] Collection support
  - [ ] Most Recently Used (MRU) support
  - [ ] Manual Save mode
- [ ] Passing UUID to SettingsModel
- [ ] Refactoring SettingsModel to move XMLSettings specific code into own class
- [ ] Refactoring singleton implementation to fit into MVVM's AppContext better