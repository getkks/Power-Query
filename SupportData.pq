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

File			: SupportData.pq @ d:\Development\M\Power-Query\SupportData.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

2020-07-29		IN086482		Added error message in case of no file path.
*/
/*
Parameters
----------

Options
	- Path 			→ Support Path. Defaults to P[Support Path].
	- File Name 	→ Partial or full name of Support File. Defaults to Support.
*/
(optional Options as record) as record => 
	let
		Selector = (Source as table, TableName as text) as table => try Table.SelectRows(
				Source{[Name = "tblSupportData"]}[Data], 
				each [Name] = TableName
			) otherwise Table.FromColumns(
				{
					Table.ColumnNames(Source{[Name = TableName]}[Data]), 
					Table.ColumnNames(Source{[Name = TableName]}[Data])
				}, 
				{"Column Name", "Output Column Name"}
			)
	in
		F[FP][ChainOperations](
				{
					{Table.SelectRows, each [Kind] = "Sheet" or [Kind] = "Text Document"}, 
					{Table.RenameColumns, {{"Sheet Name", "Name"}, {"WorkSheet", "Data"}}}, 
					{
						Table.TransformColumns, 
						{
							{
								"Data", 
								each F[FP][ChainOperations](
										{
											{Table.PromoteHeaders}, 
											{
												each Table.ReplaceValue(
														_, 
														"", 
														null, 
														Replacer.ReplaceValue, 
														Table.ColumnNames(_)
													)
											}
										}
									)(_), 
								type table
							}
						}
					}, 
					{Table.SelectColumns, {"Name", "Data"}}, 
					{
						(Source) => Table.AddColumn(
								Source, 
								"Value", 
								each F[RenameColumnsByMapping]([Data], Selector(Source, [Name]), "Column Name")
							)
					}, 
					{Record.FromTable}
				}
			)(
				F[LoadFile](
						if Record.FieldOrDefault(Options, "Path", null) = null then 
						try P[Support Path]
							 otherwise error "Path field is missing in Options parameter. P[Support Path] is not defined. Unable to load Support Data as there is no valid path."
						else 
						Options[Path], 
						[
							File Name = Record.FieldOrDefault(Options, "File Name", "Support"), 
							Recursive = true, 
							Buffer = true, 
							QuoteStyle = QuoteStyle.Csv, 
							Delimiter = ","
						]
					)
			)
