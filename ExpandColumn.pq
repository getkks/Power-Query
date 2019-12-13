(
	tblData as table,
	optional Column as text,
	optional KeepOtherColumns as logical) =>

let
	KeepOtherColumns = if KeepOtherColumns = null then true else KeepOtherColumns,
	Column = if Column = null then "WorkSheet" else Column,
	Expanded = if KeepOtherColumns then Table.ExpandTableColumn(tblData, Column, Table.ColumnNames(Table.Column(tblData,Column){0})) else Table.ExpandTableColumn(Table.SelectColumns(tblData,{Column}), Column, Table.ColumnNames(Table.Column(tblData,Column){0}))
in
	Expanded
