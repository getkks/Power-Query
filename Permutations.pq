(
	prmList as list) as table =>
let
	Ordinal = List.Numbers(1,List.Count(prmList)),
	Permutations = List.Transform(Ordinal,each F[SpecificPermutation](prmList,_)),
	#"Converted to Table" = Table.FromValue(Permutations),
	#"Expanded Permutations" = Table.ExpandTableColumn(#"Converted to Table", "Value", Table.ColumnNames(List.Last(Permutations)))
in
	#"Expanded Permutations"
