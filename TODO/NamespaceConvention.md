# Convention for Namespace
## [MVVM](../MVVM/MVVM.md)
- TODO: Clean up this mess
- MVVM
  - Infrastrcture
    - Abstract
      - IView
    - Bindings
      - CommandBindings
        - CommandButtonCommandBinding (predeclared)
        - CommandManager (module)
        - ListViewCommandBinding (predeclared)
      - PropertyBindings
        - Strategies
          - CheckBoxBindingStrategy (IBindingStrategy) (class)
          - ComboBoxBindingStrategy (IBindingStrategy) (class)
          - CommandButtonBindingStrategy (IBindingStrategy) (class)
        - CheckBoxPropertyBinding (IPropertyBinding, IHandlePropertyChanged) (class)
        - ComboBoxPropertyBinding (IPropertyBinding, IHandlePropertyChanged) (class)
        - CommandButtonPropertyBinding (IPropertyBinding, IHandlePropertyChanged) (class)
      - BindingManager
      - BindingPath (predeclared)
      - PropertyChangeNotifier
  - Common
    - Commands
      - ApplyViewModelCommand (predeclared class)
      - CancelViewCommand (predeclared class)
      - OKViewCommand (predeclared class)
    - Constants
      - TransferDirections (enum module)
  - SomeSample
    - Entities
      - MyPocoFoo
      - MyPocoBar
    - ValueConverters
      - MyPocoFooToTreeView
      - MyPocoBarToListView
    - Models
      - 💡*TTT 2.x has the entities in this folder*
      - SomeModel (class)
    - ViewModels
      - SomeViewModel (class)
    - Views
      - SomeView (IView) (userform)
    - RunSomeSample (entrypoint module)

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