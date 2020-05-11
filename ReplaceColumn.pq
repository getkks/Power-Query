/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: ReplaceColumn.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\ReplaceColumn.pq
Created			: Tuesday, 30th April 2019 5:32:41 pm

Author			: KarthikK Selvan

Change History:

2020-05-08		KarthikK Selvan		Refactored to Functional Style
*/
(Source as table,
	Column as text,
	Func as function,
	optional Type as type) =>

let
	Check = List.PositionOf(Table.ColumnNames(Source),Column),
	Index = if Check = -1 then error "Unable to Locate the Column: """ & Column & """" else Check
in
	F[FP][ChainOperations]({
		{Table.AddColumn, "Temp Column for ReplaceColumn Function", each Func(_), Type},

		{Table.RemoveColumns,{Column}},

		{F[ReorderColumnsSubset], {"Temp Column for ReplaceColumn Function"}, Index},

		{Table.RenameColumns,{{"Temp Column for ReplaceColumn Function", Column}}}
	})(Source)
/*
(Source as table,
	Column as text,
	Func as function,
	optional Type as type) =>

let
	Check = List.PositionOf(Table.ColumnNames(#"Added Custom"),Column),
	Index = if Check = -1 then error "Unable to Locate the Column: """ & Column & """" else Check,
	#"Added Custom" = Table.AddColumn(Source, "Temp Column for ReplaceColumn Function", each Func(_), Type),
	#"ReOrder Columns" = F[ReorderColumnsSubset](#"Added Custom", {"Temp Column for ReplaceColumn Function"}, Index),
	#"Removed Column" = Table.RemoveColumns(#"ReOrder Columns",{Column}),
	#"Renamed Columns" = Table.RenameColumns(#"Removed Column",{{"Temp Column for ReplaceColumn Function", Column}})
in
	#"Renamed Columns"
*/
