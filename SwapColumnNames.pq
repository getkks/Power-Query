(Source as table,
	prmColumnList as list) =>

let
	#"Column Names" = Table.ColumnNames(Source),
	sCount = List.Count(#"Column Names"),
	cCount = List.Count(prmColumnList),
	ColumnList =
		if cCount < 2 then
			error "List of Column Names should have atleast 2 Column Names for swapping."
		else if cCount > sCount then
			error "It should not exceed the Column Count (" & Text.From(sCount) & ") of the Source Table."
		else
			prmColumnList,

	#"Index of Columns" = List.Transform(ColumnList, each List.PositionOf(#"Column Names", _, Occurrence.First)),
	#"Rename Columns" = F[RenameColumnsByIndices](Source, #"Index of Columns", List.Reverse(ColumnList))
in
	#"Rename Columns"
