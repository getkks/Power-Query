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

File			: LoadBinary.pq @ d:\Development\M\Power-Query\LoadBinary.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
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
