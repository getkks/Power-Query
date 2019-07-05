let
	Source = F[FolderToTable](
		P[Data Path],
		"",
		0),

	#"Sheet to Table" = Table.TransformColumns(Source,{
		{"WorkSheet", F[SheetToTable], type table}}),

	#"Expanded WorkSheet" = F[ExpandColumn](#"Sheet to Table", "WorkSheet"),

	#"Choose and Rename Columns" = F[RenameColumnsByMapping](
		#"Expanded WorkSheet",
		S[tblColumnMapping],
		"In")
in
	#"Choose and Rename Columns"