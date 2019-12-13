(
	FolderName as text,
	optional FileName as text,
	optional SheetorTableName as any,
	optional PreFilter as function,
	optional PostFilter as function,
	optional Fold as logical) =>

let
	Source =
		try
			Folder.Files(FolderName)
		otherwise
			Folder.Files(Table.FirstValue(Excel.CurrentWorkbook(){[Name=FolderName]}[Content])),

	#"Rename File Name" = Table.RenameColumns(Source,{
			{"Name", "File Name"},
			{"Date created", "Import Date"}}),

	FilterOpenFiles = Table.SelectRows(#"Rename File Name", each
		(if FileName <> null then
			Text.Contains([File Name], FileName, Comparer.OrdinalIgnoreCase)
		else
			true)
		and
			not Text.Contains([File Name], "~$")
		and
			[Extension] <> ".tmp"
			and
				([Attributes][Kind] = "CSV File"
				or [Attributes][Kind] = "Text File"
				or [Attributes][Kind] = "Excel File")),

	Sort = Table.Sort(FilterOpenFiles ,{{"Import Date", Order.Descending}}),

	PreFiltered = Table.Buffer(
		if PreFilter<> null then
			PreFilter(Sort)
		else
			Table.FirstN(Sort,1)),

	OpenWorkBooks = Table.AddColumn(PreFiltered, "WorkBook", each
		if [Attributes][Kind] = "Excel File" then
			Excel.Workbook([Content])
		else
			null
	, type table),

	OpenCSV = Table.AddColumn(OpenWorkBooks, "CSV", each
		if [Attributes][Kind] = "CSV File" or [Attributes][Kind] = "Text File" then
			try
				Csv.Document([Content])
			otherwise
				null
		else
			null
	, type table),

	#"Expanded WorkBook" = Table.ExpandTableColumn(OpenCSV, "WorkBook",
		{"Name", "Data", "Item", "Kind", "Hidden"},
		{"Sheet Name", "Data", "Item", "Kind", "Hidden"}),

	#"Kept Range of Rows" =
		try
			if Value.Is(SheetorTableName, type number) then
				Table.Range(#"Expanded WorkBook",SheetorTableName,1)
			else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName) = 2 then
				Table.Range(#"Expanded WorkBook",SheetorTableName{0},SheetorTableName{1})
			else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName) > 2 then
					F[SelectListOfRows](#"Expanded WorkBook")
			else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName) = 1 then
						Table.Range(#"Expanded WorkBook",SheetorTableName{0},1)
			else #"Expanded WorkBook"
		otherwise
			#"Expanded WorkBook",

	#"Filtered Sheet or Table" =
		try
			if Value.Is(SheetorTableName, type text) then
				Table.SelectRows(#"Kept Range of Rows", each Text.Contains([Sheet Name],SheetorTableName, Comparer.OrdinalIgnoreCase))
			else if Value.Is(SheetorTableName, type list) then
				Table.SelectRows(#"Kept Range of Rows", each List.Contains(SheetorTableName,[Sheet Name]))
			else
				#"Kept Range of Rows"
		otherwise
			#"Kept Range of Rows",

	SortedFiles = Table.Sort(#"Filtered Sheet or Table",{
		{"Import Date", Order.Descending},
		{"File Name", Order.Ascending},
		{"Kind", Order.Ascending},
		{"Sheet Name", Order.Ascending}}),

	#"Added WorkSheet" = Table.AddColumn(SortedFiles, "WorkSheet", each
		if [Data]=null then
			[CSV]
		else
			[Data]
	, type table),

	#"Removed Other Columns" = Table.SelectColumns(#"Added WorkSheet",{"Import Date", "File Name", "Attributes", "Sheet Name", "Kind", "WorkSheet"}),
	PostFiltered =
		if PostFilter <> null then
			PostFilter(#"Removed Other Columns")
		else
			#"Removed Other Columns",

	#"Expanded Attributes" = Table.ExpandRecordColumn(PostFiltered, "Attributes", {"Kind", "Size"}, {"File Type", "Size"}),
	#"Changed Type" = Table.TransformColumnTypes(#"Expanded Attributes",{
		{"Sheet Name", type text},
		{"Kind", type text},
		{"Size", Int64.Type},
		{"File Type", type text}}),

	#"Grouped Rows" =
		if Fold <> null and Fold then
			Table.Group(#"Changed Type", {"WorkSheet"}, {{"File Attributes", each _, type table}})
		else
			#"Changed Type"
in
	#"Grouped Rows"
