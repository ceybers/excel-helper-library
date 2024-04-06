# SAP Tree To Flat
A small exercise project I did that performs basic transformation on a table of data using a list of instructions.

## Example
```
REPLACE CONTENTS OF `Stream`
	CHANGE `EEEE` TO `North Stream`
	CHANGE `EEEF` TO `East Stream`
	CHANGE `EEEG` TO `South Stream`
REPLACE CONTENTS OF `Category`
	IF `Person Responsible` IS `John Doe` THEN `Category J`
FILTER TO KEEP IF `Person Responsible` IS ONE OF
	`John Doe`
	`Jane Doe`
```