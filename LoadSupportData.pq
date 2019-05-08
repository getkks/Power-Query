(FileName as text) as record =>
let
	Source = Excel.Workbook(
		Binary.Buffer(
			File.Contents(FileName)),
		null, true),

	#"Filtered Table" = Table.SelectRows(Source, each ([Kind] = "Table")),

	Selector = 
		(TableName as text) => 
			Table.SelectRows(
				Source{[Name="tblSupportData"]}[Content], each [Name] = TableName),

	#"Choose and Rename" = Table.AddColumn(
		#"Filtered Table", "Value",
			each F[RenameColumnsByMapping](
				[Content], Selector([Name]), "Column Name")),

	#"To Record" = Record.FromTable(#"Choose and Rename")
in
	#"To Record"