(
	tbl as table,
	reorderedColumns as list,
	offset as number) as table =>

	Table.ReorderColumns(
		tbl,
		List.InsertRange(
			List.Difference(
				Table.ColumnNames(tbl),
				reorderedColumns
			),
			offset,
			reorderedColumns
		)
	)
