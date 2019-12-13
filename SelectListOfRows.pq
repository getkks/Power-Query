(Source) =>

let
	#"Added Index" = Table.AddIndexColumn(Source, "Index", 0),
	#"Filtered Rows" = Table.SelectRows(#"Added Index", each List.Contains(F[SheetorTableName],[Index])), //Replace F[SheetorTableName] with SheetorTableName if you are not using loader function to dynamically load the fuctions.
	#"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Index"})
in
	#"Removed Columns"
