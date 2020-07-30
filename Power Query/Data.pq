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

File			: FolderToTable.pq @ d:\Development\M\Power-Query\FolderToTable.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

2020-07-30		getkks		Refactored the code to Functional Style.
*/
F[FP][ChainOperations](
	{
		{Table.SelectRows, each [Kind] = "Sheet" or [Kind] = "Text Document"},
		{
				Table.TransformColumns, 
				{
					{
						"WorkSheet", 
						each F[FP][ChainOperations](
							{
								{Table.Skip, 1}, 
								{Table.PromoteHeaders}, 
								{F[RenameColumnsByMapping], S[tblSmartAPColumns], "In"}
							}
						)(_), 
						type table
					}
				}
			},
			{F[ExpandColumn], "WorkSheet"},
			{F[RenameColumnsByMapping], S[tblColumnMapping], "Out"}
	}
)(F[LoadFile](
		P[Data Path],
		[
			File Name = "",
			Recursive = true,
			Buffer = false
			//Columns = 36,//Best to avoid when unsure. When less than actual columns excess will be skipped and when higher than actual empty columns will be appended. This is applicable only for csv files. Excel Files do not depend on this parameter.
			//Encoding = 65001,//UTF8
			//QuoteStyle = QuoteStyle.None,
			//Delimiter = ","
			//Filter = (x) => x, //To avoid default filter and return all items or use any other filtering conditions
			//Extension = F[lstExtension],
			//Data Path = "", //Defaults to Folder Name parameter
			//tblColumnMapping = F[tblColumnMapping]
		]))
