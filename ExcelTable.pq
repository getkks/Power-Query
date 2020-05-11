/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: ExcelTable.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\ExcelTable.pq
Created			: Wednesday, 8th May 2019 7:23:01 pm

Author			: KarthikK Selvan

Change History:

*/
(
	TableName as text) as table =>
	Excel.CurrentWorkbook(){[Name=TableName]}[Content]
