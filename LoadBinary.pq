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
		"In2")
in
	#"Choose and Rename Columns In2"
