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

File			: tblColumnMapping.pq @ d:\Development\M\Power-Query\tblColumnMapping.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
#table(
	type table
	[
		In = text,
		In2 = text,
		Out = text,
		#"Output Column Name" = text,
		#"Type Name" = text
	],
	{
		{ "Date created" , "Import Date" , "Import Date" , "Import Date" , "DateTime"},
		{ "Name" , "File Name" , "File Name" , "File Name" , "Text"},
		{ null , "Kind" , "File Type" , "File Type" , "Text"},
		{ null , "Size" , "Size" , "Size" , "Number"},
		{ null , null , "Name" , "Sheet Name" , "Text"},
		{ null , null , "Kind" , "Kind" , "Text"},
		{ "Extension" , "Extension" , null , "Extension" , "Text"},
		{ "Content" , "Content" , null , "Content" , "Binary"},
		{ "Date accessed" , "Date accessed" , null , "Date accessed" , "DateTime"},
		{ "Date modified" , "Date modified" , null , "Date modified" , "DateTime"},
		{ null , "Content Type" , null , "Content Type" , "Text"},
		{ "Folder Path" , "Folder Path" , null , "Folder Path" , "Text"},
		{ "Attributes" , null , null , "Attributes" , null},
		{ null , "ReadOnly" , null , "ReadOnly" , "Boolean"},
		{ null , "Hidden" , null , "Hidden" , "Boolean"},
		{ null , "System" , null , "System" , "Boolean"},
		{ null , "Directory" , null , "Directory" , "Boolean"},
		{ null , "Archive" , null , "Archive" , "Boolean"},
		{ null , "Device" , null , "Device" , "Boolean"},
		{ null , "Normal" , null , "Normal" , "Boolean"},
		{ null , "Temporary" , null , "Temporary" , "Boolean"},
		{ null , "SparseFile" , null , "SparseFile" , "Boolean"},
		{ null , "ReparsePoint" , null , "ReparsePoint" , "Boolean"},
		{ null , "Compressed" , null , "Compressed" , "Boolean"},
		{ null , "Offline" , null , "Offline" , "Boolean"},
		{ null , "NotContentIndexed" , null , "NotContentIndexed" , "Boolean"},
		{ null , "Encrypted" , null , "Encrypted" , "Boolean"},
		{ null , null , "Data" , "WorkSheet" , null}})
