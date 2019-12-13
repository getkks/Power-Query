(
	Source as table,
	From as list,
	To as list) as table =>
	List.Accumulate(
		List.Zip({From, To}),
		Source,
		(state, current) =>
			Table.DuplicateColumn(
				state,
				current{0},
				current{1}))
