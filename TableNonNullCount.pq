(
	Source as any,
	optional Column as text) as number =>
let
	Source = if Value.Is(Source, type table) then Source else F[Cache](Source, FQ),
	Column = null,
	Count = List.NonNullCount(Table.Column(Source, if Column <> null then Column else List.First(Table.ColumnNames(Source))))
in
	Count
