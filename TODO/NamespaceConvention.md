# Convention for Namespace
## [MVVM](../MVVM/MVVM.md)
- MVVM
  - Infrastrcture
    - Abstract
      - IView
  - SomeSample
    - Entities
      - MyPocoFoo
      - MyPocoBar
    - ValueConverters
      - MyPocoFooToTreeView
      - MyPocoBarToListView
    - Models
      - SomeModel (class)
    - ViewModels
      - SomeViewModel (class)
    - Views
      - SomeView (IView) (userform)

## [PersistentStorage](../PersistentStorage)
- PersistentStorage
  - Abstract
    - ISettings (interface class)
    - ISettingsModel (interface class)
  - MyDocSettings
    - MyDocSettings (predeclared) (class)
  - XMLSettings
    - CustomXMLNodeHelpers (module)
    - XMLSettingsFactory (module)
    - XMLSettings (predeclared) (class)
  - SettingsModel (predeclared) (class)

## [DebugEx](../DebugEx/DebugEx.md)
- Logging
  - Abstract
    - IDebugEx (interface class)
    - IloggingProvider (interface class)
  - Model
    - DebugMessage (module)
  - Providers
    - FileLogginProvider (predeclared) (class)
    - ImmediateLoggingProvider (predeclared) (class)
  - StaticLog (module)
  - DebugEx  (predeclared) (class)

- Helpers
  - CollectionEx
    - IMostRecentlyUsed
    - MostRecentlyUsed

## ??
- Constants
- Helpers (utility, converter, control, object extender)
  - Base64, Collection, ImageList, ListObject
  - RangeToList
- Modules???