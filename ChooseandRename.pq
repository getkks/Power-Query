(
	Source as table,
	OldColumnNames as list,
	NewColumnNames as list) =>
let
	#"Reordered Columns" = F[ReorderColumnsSubset](Source, OldColumnNames, 0),
	ColumnCount=List.Count(OldColumnNames),
	#"Filtered Columns" = Table.SelectColumns(#"Reordered Columns", OldColumnNames),//List.FirstN(Table.ColumnNames(#"Reordered Columns"), ColumnCount)),
	#"Renamed Columns" = F[RenameColumnsByIndices](#"Filtered Columns", List.Numbers(0, ColumnCount),NewColumnNames)
in
	#"Renamed Columns"
