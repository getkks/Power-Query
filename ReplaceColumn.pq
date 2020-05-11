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

File			: ReplaceColumn.pq @ d:\Development\M\Power-Query\ReplaceColumn.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(Source as table,
	Column as text,
	Func as function,
	optional Type as type) =>

let
	Check = List.PositionOf(Table.ColumnNames(Source),Column),
	Index = if Check = -1 then error "Unable to Locate the Column: """ & Column & """" else Check
in
	F[FP][ChainOperations]({
		{Table.AddColumn, "Temp Column for ReplaceColumn Function", each Func(_), Type},

		{Table.RemoveColumns,{Column}},

		{F[ReorderColumnsSubset], {"Temp Column for ReplaceColumn Function"}, Index},

		{Table.RenameColumns,{{"Temp Column for ReplaceColumn Function", Column}}}
	})(Source)
