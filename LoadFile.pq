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

File			: LoadFile.pq @ d:\Development\M\Power-Query\LoadFile.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

2020-05-29		getkks		Expanded support for Zip files with more than one file. Each File name will be listed in Sheet Name column without extension
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
Encoding -> Csv File Encoding to use. UTF8 is the default encoding
QuoteStyle -> QuoteStyle for CSV File
Delimiter -> Delimiter for CSV File
Columns -> Number of Columns to process in CSV file. Do not specify column number when handling zip file with varying column sized csv files.
*/
F[FP][ChainOperations](
		{
			{each if Record.FieldOrDefault(Options, "Recursive", false) then Folder.Files(_) else Folder.Contents(_)},
			{
				Table.SelectRows,
				each (Text.Contains([Name], Record.FieldOrDefault(Options, "File Name"), Comparer.OrdinalIgnoreCase))
					and (not Text.StartsWith([Name], "~$"))
			},
			{
				Table.Join,
				{"Extension"},
				Table.FromColumns(
						{Record.FieldOrDefault(Options, "Extension", F[lstExtension])},
						type table[Extension = text]
					),
				{"Extension"},
				JoinKind.Inner,
				JoinAlgorithm.RightIndex
			},
			{Table.ExpandRecordColumn, "Attributes", {"Content Type", "Kind", "Size"}},
			{F[RenameColumnsByMapping], Record.FieldOrDefault(Options, "Column Mapping", F[tblColumnMapping]), "In2"},
			{Table.Sort, {{"Import Date", Order.Descending}}},
			{(x) => Record.FieldOrDefault(Options, "Filter", each Table.FirstN(_, 1))(x)},
			{
				F[FP][AddColumns](
						{
							{"Excel", each Text.Contains([File Type], "Excel", Comparer.OrdinalIgnoreCase), type logical},

							{"CSV", each Text.Contains([Content Type], "text", Comparer.OrdinalIgnoreCase), type logical},

							{"Zip", each Text.Contains([Content Type], "zip", Comparer.OrdinalIgnoreCase), type logical},

							{"WorkBook", (rec) =>
								if rec[Excel] then
									//Excel.Workbook(BufferedFile, false, false)
									Excel.Workbook(rec[Content], false, true)
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
											List.Transform(
												if rec[Zip] then
													Table.ToRows(
														Table.SplitColumn(
															F[LoadZip](
																rec[Content]),
															"File Name",
															Splitter.SplitTextByEachDelimiter(
																{"."},
																QuoteStyle.Csv,
																true),
															{"Name"}))
												else
													{{null, rec[Content]}},
												each {
														List.First(_), Csv.Document(List.Last(_), [
															Encoding=Record.FieldOrDefault(Options, "Encoding", 65001),
															QuoteStyle=Record.FieldOrDefault(Options, "QuoteStyle", QuoteStyle.None),
															Delimiter=Record.FieldOrDefault(Options, "Delimiter", ","),
															Columns=Record.FieldOrDefault(Options, "Columns", null)]), null, "Text Document", false})
											)
									otherwise
										null
								else
									null, type table}
						})},
            {Table.ExpandTableColumn, "WorkBook", {"Name", "Data", "Item", "Kind"}, {"Name", "Data", "Item", "Kind"}},
            {F[RenameColumnsByMapping], Record.FieldOrDefault(Options, "Column Mapping", F[tblColumnMapping]), "Out"},
            {(_) => if Record.FieldOrDefault(Options, "Buffer", false) then Table.Buffer(_) else _}
        }
    )(Record.FieldOrDefault(Options, "Data Path", FolderName))
