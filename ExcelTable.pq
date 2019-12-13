(
	TableName as text) as table =>
	Excel.CurrentWorkbook(){[Name=TableName]}[Content]
