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

File			: RenameColumnsByMapping.pq @ d:\Development\M\Power-Query\RenameColumnsByMapping.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(
	Source as table, 
	MappingTable as table, 
	FromColumn as text, 
	optional ToColumn as text, 
	optional TypeColumn as text, 
	optional Culture as text
) => 
	let
		Source = F[CleanColumnNames](Source),
		ToColumn = if ToColumn = null then "Output Column Name" else ToColumn, // otherwise "Output Column Name",
		TypeColumn = if TypeColumn = null then "Type Name" else TypeColumn,
		#"Filtered Rows" = Table.SelectRows(
				MappingTable, 
				each (Record.Field(_, FromColumn) <> null and Record.Field(_, ToColumn) <> null)
			),
		//#"Removed Other Columns" = F[ChooseandRename](#"Filtered Rows",{FromColumn, ToColumn}, {"From", "To"}),
		//#"Choose and Rename Columns" = F[ChooseandRename](Source ,#"Removed Other Columns"[From], #"Removed Other Columns"[To])
		#"Choose and Rename Columns" = F[ChooseandRename](
				Source, 
				Table.Column(#"Filtered Rows", FromColumn), 
				Table.Column(#"Filtered Rows", ToColumn)
			),
		Type = List.Contains(Table.ColumnNames(#"Filtered Rows"), TypeColumn),
		Merged = 
			if Type then 
				Table.ExpandTableColumn(
						Table.NestedJoin(
								#"Filtered Rows", 
								{TypeColumn}, 
								F[TypeMapping], 
								{"Type Name"}, 
								"Type", 
								JoinKind.LeftOuter
							), 
						"Type", 
						{"Type"}
					)
			else 
				#"Choose and Rename Columns",
		Typed = 
			if Type then 
				Table.TransformColumnTypes(
						#"Choose and Rename Columns", 
						List.Zip({Table.Column(Merged, ToColumn), Table.Column(Merged, "Type")}), 
						if Culture is null then Culture.Current else Culture
					)
			else 
				#"Choose and Rename Columns"
	in
		Typed
