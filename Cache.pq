(
	Name as text,
	optional Source as record) =>
let
	Book = Excel.CurrentWorkbook(),
	#"Transposed Table" = Table.Transpose(Book),
	#"Reversed Rows" = Table.ReverseRows(#"Transposed Table"),
	#"Promoted Headers" = Table.PromoteHeaders(#"Reversed Rows", [PromoteAllScalars=true]),
	#"To Record" = List.First(Table.ToRecords(#"Promoted Headers")),
	Query = Record.FieldOrDefault(
		if Source = null then Q else Source,
		Name,
		null),
	#"Retrieve Cache" = Record.FieldOrDefault(
		#"To Record",
		Name,
		Query),
	#"Cache or Query" = if #"Retrieve Cache" <> null and Table.RowCount(#"Retrieve Cache") > 1 then #"Retrieve Cache" else Query
in
	#"Cache or Query"