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
