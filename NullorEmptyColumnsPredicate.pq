(
	tblData as table,
	Columns as list,
	optional Or as logical) =>
let
	Source = Table.FromList(Columns),
	#"Added Custom" = Table.AddColumn(
		Source,
		"Predicate",
		each "[" & [Column1] & "]" & " <> null and [" & [Column1] & "]" & " <> """""),

	Predicate = Text.Combine(#"Added Custom"[Predicate], if Or = null or Or = true then " or " else " and "),

	Selected=Table.SelectRows(tblData, each Expression.Evaluate(Predicate, [ _=_ ]))
in
	Selected
