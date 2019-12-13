(
	Source as table,
	Indices as list,
	ColumnNamesNew as list) =>

let
	ColumnNames = Table.ColumnNames(Source),
	ColumnNamesList = List.Transform( Indices, each ColumnNames{_} ),
	ZippedList = List.Zip( { ColumnNamesList, ColumnNamesNew } ),
	#"Renamed Columns" = Table.RenameColumns( Source, ZippedList )
in
	#"Renamed Columns"
