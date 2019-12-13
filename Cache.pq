(
	Name as text,
	optional Source as record,
	optional KeyColumn as text) =>
let
	Book = Excel.CurrentWorkbook(),

	#"Transposed Table" = Table.Transpose(Book),

	#"Reversed Rows" = Table.ReverseRows(#"Transposed Table"),

	#"Promoted Headers" = Table.PromoteHeaders(#"Reversed Rows", [PromoteAllScalars=true]),

	#"To Record" = List.First(Table.ToRecords(#"Promoted Headers")),

	#"Retrieve Cache" = Record.FieldOrDefault(
		#"To Record",
		Name,
		null),

	Condition =
		if KeyColumn = null then
			#"Retrieve Cache" <> null and Table.RowCount(#"Retrieve Cache") > 1
		else
			#"Retrieve Cache" <> null and List.NonNullCount(Table.Column(#"Retrieve Cache", KeyColumn)) > 0,

	#"Cache or Query" =
		if Condition then
			#"Retrieve Cache"
		else
			Record.FieldOrDefault(
				if Source = null then Q else Source,
				Name,
				null)
in
	#"Cache or Query"
