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

File			: LoadSupportData.pq @ d:\Development\M\Power-Query\LoadSupportData.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(
	optional FilePath as text) as record =>
let
	FilePath = if FilePath = null then P[Support File] else FilePath,

	Source = Excel.Workbook(
		Binary.Buffer(
			File.Contents(FilePath)),
		null, true),

	#"Filtered Table" = Table.SelectRows(Source, each ([Kind] = "Table" and [Name] <> "tblSupportData")),

	Selector =
		(TableName as text) =>
			try Table.SelectRows(
				Source{[Name="tblSupportData"]}[Data], each [Name] = TableName)
			otherwise
				Table.FromColumns({
					Table.ColumnNames(Source{[Name=TableName]}[Data]),
					Table.ColumnNames(Source{[Name=TableName]}[Data])},
					{"Column Name", "Output Column Name"}),

	#"Choose and Rename" = Table.AddColumn(//Table.Buffer(Table.AddColumn(
		#"Filtered Table", "Value",
			each F[RenameColumnsByMapping](
				[Data], Selector([Name]), "Column Name")),//),

	#"To Record" =
		Record.FromTable(#"Choose and Rename")
in
	#"To Record"
