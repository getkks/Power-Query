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

File			: Cache.pq @ d:\Development\M\Power-Query\Cache.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

2020-05-12		KarthikK Selvan		Refactored the code to Functional Style
*/
(
	Name as text,
	optional Source as record,
	optional KeyColumn as text) =>
let
	#"Retrieve Cache" =
		Record.FieldOrDefault(
			List.First(
				Table.ToRecords(
					Table.PromoteHeaders(
						Table.ReverseRows(
							Table.Transpose(Excel.CurrentWorkbook())),
						[PromoteAllScalars=true]))),
		Name,
		null),

	Condition =
		if KeyColumn = null then
			#"Retrieve Cache" <> null and Table.RowCount(#"Retrieve Cache") > 1
		else
			#"Retrieve Cache" <> null and List.NonNullCount(Table.Column(#"Retrieve Cache", KeyColumn)) > 0,

	#"Cache or Query" =
		if Condition then
			#"Retrieve Cache"
		else
			Record.FieldOrDefault(
				if Source = null then Q else Source,
				Name,
				null)
in
	#"Cache or Query"
