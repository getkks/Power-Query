/*

Copyright 2020 - 2020 getkks, getkks@live.in

Unlicense

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http: //unlicense.org>

Project			: Power-Query @ d:\Development\M\Power-Query

File			: FolderToTable.pq @ d:\Development\M\Power-Query\FolderToTable.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
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
