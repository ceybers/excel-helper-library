# My User Defined Functions
## HashCellUDFs
- `Public Function HashCellsMD5(ParamArray Range() As Range) As Variant`
- `Public Function HashCellsSHA256(ParamArray Range() As Range) As Variant`

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