(FolderName as text, optional FileName as text, optional SheetorTableName as any, optional PreFilter as function, optional PostFilter as function) =>
let
    Source = try Table.FirstValue(Excel.CurrentWorkbook(){[Name=FolderName]}[Content]) otherwise FolderName,
    OpenFolder= Folder.Files(Source),
    Sort = Table.Sort(OpenFolder ,{{"Date created", Order.Descending}}),
    FilterOpenFiles = Table.SelectRows(Sort, each (if FileName <> null then Text.Contains([Name], FileName, Comparer.OrdinalIgnoreCase) else true) and not Text.Contains([Name], "~$") and ([Attributes][Kind] = "CSV File" or [Attributes][Kind] = "Text File" or [Attributes][Kind] = "Excel File")),
    PreFiltered = if PreFilter<> null then PreFilter(FilterOpenFiles) else Table.FirstN(FilterOpenFiles,1),
    OpenWorkBooks = Table.AddColumn(PreFiltered, "WorkBook", each if [Attributes][Kind] = "Excel File" then Excel.Workbook([Content]) else null, type table),
    OpenCSV = Table.AddColumn(OpenWorkBooks, "CSV", each if [Attributes][Kind] = "CSV File" or [Attributes][Kind] = "Text File" then try Csv.Document([Content]) otherwise null else null, type table),
    #"Expanded WorkBook" = Table.ExpandTableColumn(OpenCSV, "WorkBook", {"Name", "Data", "Item", "Kind", "Hidden"}, {"Sheet Name", "Data", "Item", "Kind", "Hidden"}),
    #"Kept Range of Rows" = if Value.Is(SheetorTableName, type number) then Table.Range(#"Expanded WorkBook",SheetorTableName,1) else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName)=2 then Table.Range(#"Expanded WorkBook",SheetorTableName{0},SheetorTableName{1}) else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName)>2 then F[SelectListOfRows](#"Expanded WorkBook") else if Value.Is(SheetorTableName, type list) and List.Count(SheetorTableName)=1 then Table.Range(#"Expanded WorkBook",SheetorTableName{0},1) else #"Expanded WorkBook",
    #"Filtered Sheet or Table" = if Value.Is(SheetorTableName, type text) then Table.SelectRows(#"Kept Range of Rows", each ([Sheet Name] = SheetorTableName)) else #"Kept Range of Rows",
    SortedFiles = Table.Sort(#"Filtered Sheet or Table",{{"Date created", Order.Descending}, {"Name", Order.Ascending}, {"Kind", Order.Ascending}, {"Sheet Name", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(SortedFiles, "WorkSheet", each if [Data]=null then [CSV] else [Data], type table),
    #"Removed Other Columns" = Table.SelectColumns(#"Added Custom",{"Date created", "Name", "Attributes", "Sheet Name", "Kind", "WorkSheet"}),
    PostFiltered = if PostFilter<> null then PostFilter(#"Removed Other Columns") else #"Removed Other Columns",
    #"Rename File Name" = Table.RenameColumns(PostFiltered,{{"Name", "File Name"}, {"Date created", "Import Date"}}),
    #"Expanded Attributes" = Table.ExpandRecordColumn(#"Rename File Name", "Attributes", {"Kind", "Size"}, {"File Type", "Size"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Attributes",{{"Sheet Name", type text}, {"Kind", type text}, {"Size", Int64.Type}, {"File Type", type text}})
in
    #"Changed Type"