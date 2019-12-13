(
	Source as table,
	KeyTable as table,
	Keys as list) as table =>
	Table.RemoveMatchingRows(
		Source,
		Table.ToRecords(KeyTable),
		Keys)
