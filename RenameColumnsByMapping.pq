(Source as table,
	MappingTable as table,
	FromColumn as text,
	optional ToColumn as text,
	optional TypeColumn as text) =>

let
	ToColumn =
		if ToColumn = null then
			"Output Column Name"
		else
			ToColumn,// otherwise "Output Column Name",
	TypeColumn  =
		if TypeColumn = null then
			"Type"
		else
			TypeColumn,
	#"Filtered Rows" = Table.SelectRows(MappingTable, each (
			Record.Field(_, FromColumn) <> null
		and
			Record.Field(_, ToColumn) <> null)),
	//#"Removed Other Columns" = F[ChooseandRename](#"Filtered Rows",{FromColumn, ToColumn}, {"From", "To"}),
	//#"Choose and Rename Columns" = F[ChooseandRename](Source ,#"Removed Other Columns"[From], #"Removed Other Columns"[To])
	#"Choose and Rename Columns" = F[ChooseandRename](Source,
		Table.Column(#"Filtered Rows", FromColumn),
		Table.Column(#"Filtered Rows", ToColumn)),

	Typed =
		if List.Contains(Table.ColumnNames(#"Filtered Rows"),TypeColumn) then
			Table.TransformColumnTypes(#"Choose and Rename Columns",
				List.Zip({
					Table.Column(#"Filtered Rows", ToColumn),
					Table.Column(#"Filtered Rows", TypeColumn)}))
		else
			#"Choose and Rename Columns"
in
	Typed