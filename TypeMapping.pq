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

File			: TypeMapping.pq @ d:\Development\M\Power-Query\TypeMapping.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
#table(
	type table
	[
		#"Type Name" = text,
		Type = type,
		#"Type Expression" = text
	],
	{
		{ "Text", type text, "text"},
		{ "String", type text, "text"},
		{ "Number", type number, "number"},
		{ "Decimal", type number, "number"},
		{ "Decimal Number", type number, "number"},
		{ "Currency", Currency.Type, "Currency.Type"},
		{ "Int64", Int64.Type, "Int64.Type"},
		{ "Whole Number", Int64.Type, "Int64.Type"},
		{ "Percentage", Percentage.Type, "Percentage.Type"},
		{ "DateTime", type datetime, "datetime"},
		{ "Date", type date, "date"},
		{ "Time", type time, "time"},
		{ "DateTimeZone", type datetimezone, "datetimezone"},
		{ "Duration", type duration, "duration"},
		{ "Logical", type logical, "logical"},
		{ "Boolean", type logical, "logical"},
		{ "Type", type type, "type"},
		{ "", type any, "any"},
		{ null, type any, "any"},
		{ "List", type list, "list"},
		{ "Record", type record, "record"},
		{ "Guid", Guid.Type, "Guid.Type"},
		{ "Binary", type binary,"binary"},
		{ "Table", type table,"table"},
		{ "Int8", Int8.Type, "Int8.Type"},
		{ "UInt8", Byte.Type, "Byte.Type"},
		{ "Int16", Int16.Type, "Int16.Type"},
		{ "Int32", Int32.Type, "Int32.Type"},
		{ "Single", Single.Type, "Single.Type"},
		{ "Double", Double.Type, "Double.Type"}})
