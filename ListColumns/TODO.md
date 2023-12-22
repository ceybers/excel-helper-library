# ListColumn Analyzer TODO
## VarType Checker
- Return Analysis on a single ListColumn
- Property of how many cells to check
    - (i.e., "Limit preview to first 1000 rows)
- Deferred update, e.g., `Update()`
- How do we return results?
    - Get(VarType enum) = value?
    - .VarType properties for each type?
    - Dictionary of `<VarType, Long>`?
    - And then we need another interface that instead of returning absolute values, returns enum of `None`, `Some`, `All`.
        - So we can easily check if VarType of a ListColumn's DataBodyRange is exclusively `vbString`.
    - Theoretically, we could use bit flags to store whether or not a ListColumn has at least one of each type.
        - If the variable has more than two flags, they are all by definition `Some`.
        - If it has exactly one flag, it is by definition `All`.

### Further Analysis
- Count of Duplicates, Count of Unique
- HasDuplicates, IsUnique
## Presentation Layer
- Has Formula
- Has Validation
- Has ConditionalFormats
- Is Protected (andLocked, andUnlocked)
- HasSpecialCells(xlCellType)
    ??? I think this was a faster way of checking for blanks
    Remember we only have to check for at least one blank, we don't need to evaluate every cell.
### Thoughts
- We probably need a conversion method to get "only text and only distinct (removed duplicates)" before we try analysing two columns for mapping.
## Two ListColumns Set Theory
- A is subset of B, B is subset of A
- Outer join, inner join, anti-join
- The actual values, the counts, and the percentages.
- Deferred evaluate with first `n` row limits.
## Two ListColumns Src to Dst
- How do we handle this?
    - Three lists (orig, cur src, cur dst)?
    - If we are trying to model in the spirit of CollectionView, this would be necessary.
        - Would we then hold a Queue of ICommand with Undo and Redo to move items back and forth?
    - Sounds like a very heavy-weight model.
    - All we wanted to do was compare elements between two sets.