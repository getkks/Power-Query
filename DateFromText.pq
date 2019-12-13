(
	Source as text,
	optional Positions as list/*list meta [
		Documentation.FieldCaption = "Positions",
		Documentation.FieldDescription = "List of positions for date splitter.",
		Documentation.SampleValues =  {"{0,2,4}"}]*/,
	optional Combiner as function) as date=>
let
	Splitter = Splitter.SplitTextByPositions(
		if Positions is null then
			{0,2,4}
		else
			Positions),

	Splited = Splitter(Source),

	Split = List.Transform(Splited, Int16.From),

	Date =
		if Combiner is null then
			#date(Split{2},Split{0},Split{1})
		else
			Combiner(Split)
in
	Date
