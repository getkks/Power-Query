(Search as nullable any,
	KeyValue as table) =>

let
	Columns = Table.ToColumns(KeyValue),
	Position = List.PositionOf(Columns{0},Search, Occurrence.First),
	Replaced = if Position = -1 then Search else Columns{1}{Position}
in
	Replaced