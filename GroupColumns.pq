(
	Source as table,
	Delimiter as text,
	optional IndexType as type,
	optional Reverse as logical) =>
let
	ColumnNames = Table.ColumnNames(Source),
	//#"Removed Top Items" = List.Skip(ColumnNames,11),

	#"Converted to Table" = Table.FromColumns(
		{ColumnNames,ColumnNames},
		{"Column Names", "Columns"}),

	#"Split Column by Delimiter" = Table.SplitColumn(
		#"Converted to Table",
		"Columns",
		Splitter.SplitTextByEachDelimiter(
			{Delimiter}, QuoteStyle.None, true),
			{"Column", "Columns Index"}),

	#"Grouped Rows" = Table.Group(
		#"Split Column by Delimiter",
		{"Columns Index"}, {
			{"Consolidated", each _, type table [Column Names=text, Column=text, Columns Index=text]}}),

	#"Changed Type" = Table.TransformColumnTypes(#"Grouped Rows",{
		{"Columns Index",
			if IndexType <> null then
				IndexType
			else
				Int32.Type}}),

	#"Sorted Rows" = Table.Sort(#"Changed Type",{
		{"Columns Index",
			if Reverse = null or not Reverse then
				Order.Ascending
			else
				Order.Descending}})
in
	#"Sorted Rows"
