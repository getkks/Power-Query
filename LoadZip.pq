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

File			: LoadZip.pq @ d:\Development\M\Power-Query\LoadZip.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
(ZIPFileContent as binary) as table =>
let
	Header = BinaryFormat.Record([
		Signature			= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
		Version 			= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		Flags				= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		Compression 		= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		ModTime 			= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		ModDate 			= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		CRC32				= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
		CompressedSize		= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
		UncompressedSize	= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
		FileNameLen 		= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
		ExtraFieldLen		= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian)]),

	FileEntry = BinaryFormat.Choice(Header, each
		if _[Signature] <> 0x4034B50 then
			BinaryFormat.Null
		else
			BinaryFormat.Record([
				Header			 = _,
				FileName		 = BinaryFormat.Text(_[FileNameLen]),
				ExtraField		 = BinaryFormat.Text(_[ExtraFieldLen]),
				UncompressedData = BinaryFormat.Transform(BinaryFormat.Binary(_[CompressedSize]),(x) => try Binary.Decompress(x, Compression.Deflate) otherwise null)
				]),
		type binary),

	ZipFormat = BinaryFormat.List(FileEntry, each _<> null),

	Entries = List.Transform(
			List.RemoveLastN(ZipFormat(Binary.Buffer(ZIPFileContent)), 1),
		(e) => [File Name = e[FileName], Content = e[UncompressedData] ]
	)
in
	Table.FromRecords(Entries)
