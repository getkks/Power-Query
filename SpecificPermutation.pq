(Source as list, SetNumber as number) as table =>
let
	// number of combinations
	NumberOfCombinations = Number.Combinations(List.Count(Source), SetNumber),

	/* create SetNumber - 1 Lists, skipping one progressive items from Source
	   For a "A B C D E" list for which we want to generate combinations of 3, return 3 lists of:
	   A  B C D E
	   B C D E
	   C D E
	*/
	PartitionList = List.Buffer(List.Accumulate(
								  {0..SetNumber - 1},
								  {},
								  (state, current) => state & {List.Skip(Source, current)}
							   )),

	// item count of the last item of PartitionList
	LastListLength = List.Count(List.Last(PartitionList)),

	// generate list of 1's of length equal to setNumber
	FirstListItems = List.Repeat({1}, SetNumber),

	// set number items positions,
	SetNumberItemsPosition = List.Buffer({0..SetNumber - 1}),

	// for each item in Sequence list, take the correspondend item of Partition list
	AssignItems = (Sequence as list) as list =>
			List.Accumulate(
					SetNumberItemsPosition,
					{},
					(state, current) => state & {PartitionList{current}{Sequence{current}-1}}
				),


	GenerateNewNumList = (PrevNumList as list) =>
		List.Accumulate(
				SetNumberItemsPosition,
				{},
				(state, current) =>
								   let
									 Res = if
											  current = SetNumber - 1
										   then
											  if PrevNumList{current} = LastListLength then state{current - 1} else PrevNumList{current} + 1
										   else
											  if
												  PrevNumList{current + 1} = LastListLength and PrevNumList{current} < LastListLength
											  then
												  PrevNumList{current} + 1
											  else
												  if PrevNumList{current} = LastListLength then state{current - 1} else PrevNumList{current}
									in
										state & {Res}
		   ),

	Result = List.Generate(
							()=>[NumList = FirstListItems, Combin = AssignItems(NumList), counter = 1],
							each [counter] <= NumberOfCombinations,
							each [NumList = GenerateNewNumList([NumList]), Combin = AssignItems(NumList), counter = [counter] + 1],
							each [Combin]
						 ),

	FinalTable = Table.FromRows(Result)

in
	FinalTable
