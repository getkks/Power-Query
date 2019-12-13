(Source as table,
	MappingTable as table,
	FromColumn as text,
	optional ToColumn as text) =>

let
	ToColumn = if ToColumn = null then "Output Column Name" else ToColumn,// otherwise "Output Column Name",
	#"Filtered Rows" = Table.SelectRows(MappingTable, each (Record.Field(_, FromColumn) <> null and Record.Field(_, ToColumn) <> null)),
	//#"Removed Other Columns" = F[ChooseandRename](#"Filtered Rows",{FromColumn, ToColumn}, {"From", "To"}),
	//#"Choose and Rename Columns" = F[ChooseandRename](Source ,#"Removed Other Columns"[From], #"Removed Other Columns"[To])
	#"Choose and Rename Columns" = F[ChooseandRename](Source ,Table.Column(#"Filtered Rows",FromColumn), Table.Column(#"Filtered Rows",ToColumn))
in
	#"Choose and Rename Columns"
