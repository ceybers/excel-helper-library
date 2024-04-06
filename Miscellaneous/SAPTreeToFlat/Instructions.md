LOAD FROM `Input`
TAKE COLUMNS
	`Level`
	`Project object` AS `Project Description`
	`Projektelm` AS `Project Number`
	`Budget`
	`Act. Costs` as `Actualised`
	`TtlCstComm` as `Committed`
	`Pers.resp.` AS `Person Responsible`
	`Erl. Start` as `Start Date`
	`Status`
ADD COLUMN `L5` CONTAINING LAST `Project Number` WHERE `Level` IS `05` 
FILTER TO KEEP NOTHING
FILTER TO KEEP IF `Person Responsible` IS ONE OF
	`John Doe`
	`Jane Doe`
	`John Smith`
	`Jane Smith`
FILTER TO KEEP IF `L5` IS ONE OF
	`AA-BBB-CCCC-DDD-EEEE-FFF`
	`AA-BBB-CCCC-DDF-EEEG-FFF`
FILTER TO REMOVE IF `Level` IS ONE OF
	`03`
	`04`
	`05`
REMOVE COLUMN `Level`
APPLY FILTER
ADD COLUMNS BY SPLITTING `L5` EVERY `-`
	`4` AS `Division`
	`5` AS `Stream`
	`6` AS `Category`
REMOVE COLUMN `L5`
REPLACE CONTENTS OF `Division`
	CHANGE `DDD` TO `Red Division`
	CHANGE `DDE` TO `Green Division`
	CHANGE `DDF` TO `Blue Division`
REPLACE CONTENTS OF `Stream`
	CHANGE `EEEE` TO `North Stream`
	CHANGE `EEEF` TO `East Stream`
	CHANGE `EEEG` TO `South Stream`
REPLACE CONTENTS OF `Category`
	IF `Person Responsible` IS `John Doe` THEN `Category J`
SAVE TO `Output`
	`Division`
	`Stream`
	`Category`
	`Project Number`
	`Project Description`
	`Budget`
	`Actualised`
	`Committed`
	`Start Date` IN COLUMN `M`