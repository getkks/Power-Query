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

Project        : Power-Query @ d:\Development\M\Power-Query

File           : FilterExceptions.pq @ d:\Development\M\Power-Query\FilterExceptions.pq
File Created   : Saturday, 9th May 2020 3:20:38 pm

Author         : getkks (getkks@live.in)

Change History:

 */
(
	Source as table,
	Exceptions as table,
	Columns as any,
	optional ToColumns as any,
	optional Filter as logical) =>
let
//	  Error = if not (Value.Is(Columns, type list) or Value.Is(Columns, type text)) then error "Function FilterExceptions: => Columns should be either list of column names or a single column name (text)." else null,
	ColumnList = Value.Is(Columns, type list),
	JoinType = if Filter <> null and not Filter then JoinKind.Inner else JoinKind.LeftAnti,
	toColumns = if ToColumns = null then Columns else ToColumns,
//	  MergeException = if Error <> null then Error else Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	MergeException =  Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	#"Removed Columns" = Table.RemoveColumns(MergeException,{"tblExceptions"})
in
	#"Removed Columns"

/*let
	Source = Data3,
	Exceptions = tblExceptions,
	Columns = "Account",
	ToColumns = "Account",
	Filter = false,
	Error = if not (Value.Is(Columns, type list) or Value.Is(Columns, type text)) then error "Function FilterExceptions: => Columns should be either list of column names or a single column name (text)." else null,
	ColumnList = Value.Is(Columns, type list),
	JoinType = if Filter <> null and not Filter then JoinKind.Inner else JoinKind.LeftAnti,
	toColumns = if ToColumns = null then Columns else ToColumns,
	MergeException = if Error <> null then Error else Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	#"Removed Columns" = Table.RemoveColumns(MergeException,{"tblExceptions"})
in
	#"Removed Columns"
*/
