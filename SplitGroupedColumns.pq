(
	Source as record,
	GroupedColumns as table) =>
let
	Values = Table.AddColumn(
		GroupedColumns,
		"Values", each
		Record.SelectFields(Source, [Consolidated][Column Names])),

	Names = Table.AddColumn(
		Values,
		"Record",each
		Record.RenameFields([Values],List.Zip({[Consolidated][Column Names],[Consolidated][Column]}))),

	Result = Table.FromRecords(
		Names[Record],
		Record.FieldNames(Names[Record]{0}))
in
	Result
