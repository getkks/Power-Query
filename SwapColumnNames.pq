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

Project			: Power Query Library @ d:\Development\M\Power-Query

File			: SwapColumnNames.pq @ d:\Development\M\Power-Query\SwapColumnNames.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(Source as table,
	prmColumnList as list) =>

let
	#"Column Names" = Table.ColumnNames(Source),
	sCount = List.Count(#"Column Names"),
	cCount = List.Count(prmColumnList),
	ColumnList =
		if cCount < 2 then
			error "List of Column Names should have atleast 2 Column Names for swapping."
		else if cCount > sCount then
			error "It should not exceed the Column Count (" & Text.From(sCount) & ") of the Source Table."
		else
			prmColumnList,

	#"Index of Columns" = List.Transform(ColumnList, each List.PositionOf(#"Column Names", _, Occurrence.First)),
	#"Rename Columns" = F[RenameColumnsByIndices](Source, #"Index of Columns", List.Reverse(ColumnList))
in
	#"Rename Columns"
