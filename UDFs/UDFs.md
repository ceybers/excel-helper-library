# My User Defined Functions
## HashCellUDFs
### Abstract
Returns the hash of the values in a given range.

### Description
Computes the hash of each Cell in a Range, taking into consideration its relative position in the Range. If multiple Ranges are passed as parameters, we take into consideration their order.

For each Range passed, a hash is computed for each Cell. Each of the hashes are followed by the ASCII Unit Separator character (`0x1F`). At then end of each row, the Record Separator character (`0x1E`) is appended. At the end of each Range, the Group Separator character (`0x1D`) is appended. This delimited string of hashes and control characters is then hashed once more to return one single hash for the Range.

We then concatenate the hashes for each Range (even if only one Range was passed), and then compute the hash of that concatenated string.

The purpose of the delimiters is to ensure the cells are in the same relative position in the Range. If we simply concatenate all the hashes, a 3×2 table and a 2×3 table with the same values stored from left-to-right then top-to-bottom would result in the same hash as we flatten the 2D array of cells into a 1D list of hashes.

Since the Range parameters are passed as a 1-dimensional list, we do not need to use delimiters to ensure that a different hash is returned if they are passed in a different order.

Finally, if the function is called on a single Cell and the cell is blank, the function returns the hash of the `NUL` ASCII character `0x00`. 

### Performance
- 10'000 cells using SHA1 algorithm: 0.313 seconds
- 10'000 cells using SHA256 algorithm: 0.344 seconds
- 10'000 cells using MD5 algorithm: 0.281 seconds
- 100'000 cells scaled to 10× as long, and 1'000'000 cells to 100×.

### User Defined Functions
- `HashCellsMD5(range1, range2...)`
- `HashCellsSHA1(range1, range2...)`
- `HashCellsSHA256(range1, range2...)`

## Regular Expressions
- `Regex(Value, Expression, Optional OutputPattern)`
### Examples:
```vb
? Regex("foobar", "foo")
foo

? Regex("foo", "bar")
FALSE

? Regex("foobar", "^\w{4}")
foob

? Regex("foobar", "(\w{2})(\w{2}).*", "$2")
ob
```