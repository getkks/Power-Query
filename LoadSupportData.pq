(
	optional FilePath as text) as record =>
let
	FilePath = if FilePath = null then P[Support File] else FilePath,

	Source = Excel.Workbook(
		Binary.Buffer(
			File.Contents(FilePath)),
		null, true),

	#"Filtered Table" = Table.SelectRows(Source, each ([Kind] = "Table" and [Name] <> "tblSupportData")),

	Selector =
		(TableName as text) =>
			try Table.SelectRows(
				Source{[Name="tblSupportData"]}[Data], each [Name] = TableName)
			otherwise
				Table.FromColumns({
					Table.ColumnNames(Source{[Name=TableName]}[Data]),
					Table.ColumnNames(Source{[Name=TableName]}[Data])},
					{"Column Name", "Output Column Name"}),

	#"Choose and Rename" = Table.AddColumn(//Table.Buffer(Table.AddColumn(
		#"Filtered Table", "Value",
			each F[RenameColumnsByMapping](
				[Data], Selector([Name]), "Column Name")),//),

	#"To Record" =
		Record.FromTable(#"Choose and Rename")
in
	#"To Record"