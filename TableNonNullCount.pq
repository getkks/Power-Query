/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: TableNonNullCount.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\TableNonNullCount.pq
Created			: Tuesday, 26th November 2019 1:16:21 pm

Author			: KarthikK Selvan

Change History:

*/
(
	Source as any,
	optional Column as text) as number =>
let
	Source = if Value.Is(Source, type table) then Source else F[Cache](Source, FQ),
in
	List.NonNullCount(Table.Column(Source, if Column <> null then Column else List.First(Table.ColumnNames(Source))))
