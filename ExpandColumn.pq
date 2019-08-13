(
	tblData as table,
	optional Column as text,
	optional KeepOtherColumns as logical) =>

let
	KeepOtherColumns = if KeepOtherColumns = null then true else KeepOtherColumns,
	Column = if Column = null then "WorkSheet" else Column,
	ColumnNames = Table.ColumnNames(
		List.First(
			Table.Column(
				tblData,
				Column))),
	Working =
		if KeepOtherColumns then
			tblData
		else
			Table.SelectColumns(
				tblData,
				{Column}),

	Expanded = Table.ExpandTableColumn(Working, Column, ColumnNames)
in
	Expanded