(
	optional FolderName as text,
	optional Options as record
	) as table =>
/*
Parameters
----------
FolderName -> Folder path for traversing.
Data Path -> Default folder path for traversing
File Name -> Full or Partial Name of the file to locate.
Recursive -> Default value is false. True will force the folder to be traversed recursively (all sub directories).
tblColumnMapping -> Mapping for Rename operations. Consists of three stages.
Extension -> List to filter into required file extensions.
Filter -> Filters the Folder Contents using a function which operates on table. This is post stage 1 rename and sorting. By default, recent file will be selected
Buffer -> Enables buffer before opening the file contents. By default, the file contents will not be buffered
Encoding -> Csv File Enconding to use. UTF8 is the default encoding
QuoteStyle -> QuoteStyle for CSV File
Delimiter -> Delimiter for CSV File
Columns -> Number of Columns to process in CSV file.
*/
let
	Source = Record.FieldOrDefault(Options, "Data Path", FolderName),

	OpenFolder =
		if Record.FieldOrDefault(Options, "Recursive", false) then
			Folder.Files(Source)
		else
			Folder.Contents(Source),

	#"Choose and Rename Columns In" = F[RenameColumnsByMapping](
		OpenFolder,
		Record.FieldOrDefault(
			Options,
			"Column Mapping",
			F[tblColumnMapping]),
		"In"),

	FilterOpenFiles = Table.SelectRows(#"Choose and Rename Columns In", each
			(Text.Contains(
				[File Name],
				Record.FieldOrDefault(Options, "File Name"),
				Comparer.OrdinalIgnoreCase))
		and
			(not Text.StartsWith([File Name], "~$"))),

	#"Filtered Known Extensions" = Table.FuzzyJoin(
		FilterOpenFiles, {"Extension"},
		Table.FromColumns({
			Record.FieldOrDefault(Options, "Extension", F[lstExtension])},
			type table[Extension = text]), {"Extension"},
		JoinKind.Inner, [
			NumberOfMatches = 1,
			Threshold = 1,
			ConcurrentRequests = 4]),
/*
	#"Filtered Known Extensions" = Table.Join(
		FilterOpenFiles, {"Extension"},
		Table.FromColumns({
			Record.FieldOrDefault(Options, "Extension", F[lstExtension])},
			type table[Extension = text]), {"Extension"},
		JoinKind.Inner),
*/
	#"Sort by Import Date" = Table.Sort(#"Filtered Known Extensions",{{"Import Date", Order.Descending}}),

	#"Apply PreFilter" = if Record.FieldOrDefault(Options, "Filter") <> null then
			Record.FieldOrDefault(Options, "Filter")(#"Sort by Import Date")
		else
			Table.FirstN(#"Sort by Import Date",1),

	#"Expanded Attributes" = Table.ExpandRecordColumn(#"Apply PreFilter", "Attributes", {"Content Type", "Kind", "Size", "ReadOnly", "Hidden", "System", "Directory", "Archive", "Device", "Normal", "Temporary", "SparseFile", "ReparsePoint", "Compressed", "Offline", "NotContentIndexed", "Encrypted"}),

	#"Choose and Rename Columns In2" = F[RenameColumnsByMapping](
		#"Expanded Attributes",
		Record.FieldOrDefault(
			Options,
			"Column Mapping",
			F[tblColumnMapping]),
		"In2"),

	#"Added is Excel" = Table.AddColumn(#"Choose and Rename Columns In2", "Excel", each
		Text.Contains([File Type], "Excel", Comparer.OrdinalIgnoreCase),
		type logical),

	#"Added is CSV" = Table.AddColumn(#"Added is Excel", "CSV", each
			Text.Contains([Content Type], "text", Comparer.OrdinalIgnoreCase),
		type logical),
    #"Added Zip" = Table.AddColumn(#"Added is CSV", "Zip", each
			Text.Contains([Content Type], "zip", Comparer.OrdinalIgnoreCase),
		type logical),

	ToBuffer = Record.FieldOrDefault(Options, "Buffer", false),

	#"Load WorkBook" = Table.AddColumn(#"Added Zip", "WorkBook", (rec) =>
			let
				BufferedFile = if ToBuffer then Binary.Buffer(rec[Content]) else rec[Content],
				LoadedFile =
					if rec[Excel] then
						Excel.Workbook(BufferedFile, false, false)
					else if rec[CSV] or rec[Zip] then
						try
							#table(
								type table
								[
									Name = text,
									Data = table,
									Item = text,
									Kind = text,
									Hidden = logical
								],
								{
									{null, Csv.Document(if rec[Zip] then List.First(F[LoadZip](BufferedFile)[Content]) else BufferedFile, [
										Encoding=Record.FieldOrDefault(Options, "Encoding", 65001),
										QuoteStyle=Record.FieldOrDefault(Options, "QuoteStyle", QuoteStyle.None),
										Delimiter=Record.FieldOrDefault(Options, "Delimiter", ","),
										Columns=Record.FieldOrDefault(Options, "Columns", null)]), null, "Text Document", false}
								})
						otherwise
							null
					else
						null
			in
				LoadedFile,
		type table),

	#"Expanded WorkBook" = Table.ExpandTableColumn(#"Load WorkBook", "WorkBook", {"Name", "Data", "Item", "Kind", "Hidden"}, {"Name", "Data", "Item", "Kind", "Hidden Sheet"}),

	#"Choose and Rename Columns Out" = F[RenameColumnsByMapping](
		#"Expanded WorkBook",
		Record.FieldOrDefault(
			Options,
			"Column Mapping",
			F[tblColumnMapping]),
		"Out"),

	Result = if ToBuffer then Table.Buffer(#"Choose and Rename Columns Out") else #"Choose and Rename Columns Out"
in
	Result
