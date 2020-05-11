/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: SplitGroupedColumns.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\SplitGroupedColumns.pq
Created			: Thursday, 14th November 2019 2:22:58 pm

Author			: KarthikK Selvan

Change History:

*/
(
	Source as record,
	GroupedColumns as table) =>
let
	Values = Table.AddColumn(
		GroupedColumns,
		"Values", each
		Record.SelectFields(Source, [Consolidated][Column Names])),

	Names = Table.AddColumn(
		Values,
		"Record",each
		Record.RenameFields([Values],List.Zip({[Consolidated][Column Names],[Consolidated][Column]}))),

	Result = Table.FromRecords(
		Names[Record],
		Record.FieldNames(Names[Record]{0}))
in
	Result
