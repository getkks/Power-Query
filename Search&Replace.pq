(
	Source as table,
	KeyValue as table,
	SearchColumn as nullable any) =>

let
	Key = Table.ColumnNames(KeyValue),
	#"Merged Queries" = Table.NestedJoin(Source,{SearchColumn},KeyValue,{Key{0}},"Search & Replace Temp Column 1",JoinKind.LeftOuter),
	#"Expanded Results" = Table.ExpandTableColumn(#"Merged Queries", "Search & Replace Temp Column 1", {Key{1}}, {"Search & Replace Temp Column 1"}),
	#"Added Custom" = F[ReplaceColumn](#"Expanded Results", SearchColumn, each if [#"Search & Replace Temp Column 1"] = null then Record.Field(_,SearchColumn) else [#"Search & Replace Temp Column 1"]),
	#"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Search & Replace Temp Column 1"})
in
	#"Removed Columns"
