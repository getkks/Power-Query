let
	Source = F[LoadFile](
		P[Data Path],
		[
			File Name = "",
			Recursive = true,
			Buffer = false
			//Columns = 36,
			//Encoding = 65001,//UTF8
			//QuoteStyle = QuoteStyle.None,
			//Delimiter = ","
			//Filter = (x) => x, //To avoid default filter
			//Extension = F[lstExtension],
			//Data Path = "", //Defaults to Folder Name parameter
			//tblColumnMapping = F[tblColumnMapping]
		]),

	#"Filtered Rows" = Table.SelectRows(Source, each
			[Kind] = "Sheet" /*For Excel Files. Helps avoid todefined names when filter is enabled.*/
		or
			[Kind] = "Text Document" /*For CSV file and Zip File containing CSV File*/),

	#"Sheet to Table" = Table.TransformColumns(#"Filtered Rows",{
		{"WorkSheet", F[SheetToTable], type table}}),

	#"Expanded WorkSheet" = F[ExpandColumn](#"Sheet to Table", "WorkSheet"),

	#"Choose and Rename Columns" = F[RenameColumnsByMapping](
		#"Expanded WorkSheet",
		S[tblColumnMapping],
		"In")
in
	#"Choose and Rename Columns"
