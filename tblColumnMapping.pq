/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: tblColumnMapping.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\tblColumnMapping.pq
Created			: Wednesday, 10th July 2019 6:14:32 pm

Author			: KarthikK Selvan

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
		{ "Date created" , "Date created" , "Import Date" , "Import Date" , "DateTime"},
		{ "Name" , "Name" , "File Name" , "File Name" , "Text"},
		{ null , "Kind" , "File Type" , "File Type" , "Text"},
		{ null , "Size" , "Size" , "Size" , "Number"},
		{ null , null , "Name" , "Sheet Name" , "Text"},
		{ null , null , "Kind" , "Kind" , "Text"},
		{ "Extension" , "Extension" , null , "Extension" , "Text"},
		{ "Content" , "Content" , null , "Content" , "Binary"},
		{ "Attributes" , null , null , "Attributes" , null},
		{ null , "Content Type" , null , "Content Type" , "Text"},
		{ null , null , "Data" , "WorkSheet" , null}})
/*
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
*/
/*
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
		{ "Attributes" , null , null , "Attributes" , null},
		{ null , "Content Type" , null , "Content Type" , "Text"},
		{ null , null , "Data" , "WorkSheet" , null}})
*/
