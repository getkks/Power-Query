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

File			: PrioritizeValues.pq		@ d:\Development\M\Power-Query\PrioritizeValues.pq
Created			: Monday, 6th September 2021 2:29:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
( tbl as table, sourceColumns as list, preferredValuesTable as table, matchColumns as list ) =>
	let
		columnNames 		 = List.Buffer ( Table.ColumnNames ( preferredValuesTable ) ),
		preferredValuesTable = Table.Buffer ( preferredValuesTable )
	in
		Table.Combine (
			{
				Table.RemoveColumns (
					Table.Join ( tbl, sourceColumns, preferredValuesTable, matchColumns, JoinKind.LeftAnti ),
					columnNames
				),
				Table.RemoveColumns (
					Table.Join ( tbl, sourceColumns, preferredValuesTable, matchColumns, JoinKind.Inner, JoinAlgorithm.RightIndex ),
					columnNames
				)
			}
		)
/*
( tbl as table, sourceColumns as list, preferredValuesTable as table, matchColumns as list ) =>
	Table.RemoveColumns (
		Table.Combine (
			{
				Table.Join ( tbl, sourceColumns, preferredValuesTable, matchColumns, JoinKind.LeftAnti ),
				Table.Join ( tbl, sourceColumns, preferredValuesTable, matchColumns, JoinKind.Inner )
			}
		),
		Table.ColumnNames ( preferredValuesTable )
	)
*/