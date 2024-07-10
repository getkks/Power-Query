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

File			: CodeLoader.pq		@ => d:\Development\M\Power-Query\CodeLoader.pq
Created			: Friday, 30th April 2021 3:13:28 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(Path as text) =>
	F[FP][ChainOperations]
		(
			{
				{Table.SelectRows, each Text.Lower([Extension]) = ".vba"},
				{
					F[FP][AddColumns]
						(
							{
								{"Query", each Text.Start([Name], Text.Length([Name]) - Text.Length([Extension])), type text},
								{
									"Value",
									each [Folder Path] & [Name]
								}
							}
						)
				},
				{F[RenameColumnsByMapping], F[DefaultParameter][tblQueryColumnMapping], "In"},
				{Record.FromTable}
			}
		)
		(Folder.Files(Path, [PreserveLastAccessTimes=true]))
