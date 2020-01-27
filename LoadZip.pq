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
			List.RemoveLastN( ZipFormat(ZIPFileContent), 1),
		(e) => [File Name = e[FileName], Content = e[UncompressedData] ]
	)
in
	Table.FromRecords(Entries)