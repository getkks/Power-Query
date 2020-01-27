(Source as table,
	MappingTable as table,
	FromColumn as text,
	optional ToColumn as text,
	optional TypeColumn as text,
	optional Culture as text) =>

let
	Source = F[CleanColumnNames](Source),

	ToColumn =
		if ToColumn = null then
			"Output Column Name"
		else
			ToColumn,// otherwise "Output Column Name",

	TypeColumn	=
		if TypeColumn = null then
			"Type Name"
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

	Type = List.Contains(Table.ColumnNames(#"Filtered Rows"),TypeColumn),

	Merged =
		if Type then
			Table.ExpandTableColumn(
				Table.NestedJoin(
					#"Filtered Rows", {TypeColumn},
					F[TypeMapping], {"Type Name"},
					"Type",
					JoinKind.LeftOuter),
				"Type", {"Type"})
		else
			#"Choose and Rename Columns",

	Typed =
		if Type then
			Table.TransformColumnTypes(#"Choose and Rename Columns",
				List.Zip({
					Table.Column(Merged, ToColumn),
					Table.Column(Merged, "Type")}),
				if Culture is null then Culture.Current else Culture)
		else
			#"Choose and Rename Columns"
in
	Typed