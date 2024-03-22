# Lambdas
## `TLOOKUP`
```
=LAMBDA(Table,KeyCol,Key,Field,XLOOKUP(Key,INDIRECT(Table&"["&KeyCol&"]"),INDIRECT(Table&"["&Field&"]")))

' e.g.: TLOOKUP("Table1","col1",F49:F51,"col3")
```

## `GetTable`
GetRow returns the entire row, seems wasteful. If we only want one column in that row.
```
GetColumnByName =LAMBDA(T,N,INDIRECT(T&"["&N&"]"))
GetRow =LAMBDA(T,K,OFFSET(INDIRECT(T),MATCH(K,FirstColumn(T),0)-1,,1,))

' e.g., Specific hard-coded lambda:
GetTable2 =LAMBDA(K,F,GetRow("Table2",K) GetColumnByName("Table2",F))
```

## Calendars
Calendar table sequence generators. Weeks in Year takes a day offset to allow for choosing a specific starting date (e.g., Weeks per year, starting on Tuesdays).
```
MonthsInYear =LAMBDA(Y,DATE(Y,SEQUENCE(12),1))
DaysInYear =LAMBDA(Y,SEQUENCE(DATE(Y+1,1,1)-DATE(Y,1,1),1,DATE(Y,1,1)))
WeeksInYear =LAMBDA(Y,D,DATE(Y,1,D)+(SEQUENCE(52,,0)*7))(2024,3)
```

## `CountIfMonth`
Alternative to `COUNTIFS`. Given an array of dates, returns the count of rows where the date is in the same month as the criteria.
```
=LAMBDA(D, M, COUNTIFS(D, ">=" & EOMONTH(M+0, -1)+1, D, "<=" & EOMONTH(M+0, 0)))

' e.g.: =CountIfMonth(Table1[Date], ThisMonth)
```

## `FilterByMonth`
Helper lambda for `FILTER` that includes all rows where the date in the filter column is in the same month as the criteria.
```
=LAMBDA(D,M,EOMONTH(D+0,0)=EOMONTH(M,0))

` e.g.: =FILTER(Table1[ID], FilterByMonth(Table1[PS], ThisMonth))
```

## Before, Between, After a Date Range
Takes an input of a Start date, an End date, and a Day to test. 

Returns the appropriate string based on whether the criteria is Before, Between or After the range.
```
BetweenDates =LAMBDA(Start,End,Day,IFS(MIN(Start,Day)<Start,"Before",MAX(End,Day)>End,"After",TRUE,"Between"))
```