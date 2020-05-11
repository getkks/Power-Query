/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: CleanColumnNames.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\CleanColumnNames.pq
Created			: Friday, 12th July 2019 3:56:47 pm

Author			: KarthikK Selvan

Change History:

2020-05-08		KarthikK Selvan		Refactored to Functional Style
*/
(Source as table) as table =>
let
	#"Column Names" = Table.ColumnNames(Source)
in
	Table.RenameColumns(Source,
		List.Zip({#"Column Names",
			List.Transform(
				#"Column Names", each
					Text.Clean(
						Text.Trim(_)))}))
