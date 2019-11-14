let
	Source = F[LoadFile](
		P[Data Path],
		[
			File Name = "",
			Recursive = true
		]),

	#"Filtered Rows" = Table.SelectRows(Source, each [Kind] = "Sheet" or [Kind] = "Text Document"),

	#"Sheet to Table" = Table.TransformColumns(#"Filtered Rows",{
		{"WorkSheet", F[SheetToTable], type table}}),

	#"Expanded WorkSheet" = F[ExpandColumn](#"Sheet to Table", "WorkSheet"),

	#"Choose and Rename Columns" = F[RenameColumnsByMapping](
		#"Expanded WorkSheet",
		S[tblColumnMapping],
		"In")
in
	#"Choose and Rename Columns"