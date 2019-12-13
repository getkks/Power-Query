(
	Source as table,
	Exceptions as table,
	Column as text,
	optional ToColumn as text,
	optional Filter as logical) =>
let
	toColumn = if ToColumn = null then Column else ToColumn,
	Selection = Table.SelectRows(Source, each if Filter = null or Filter then not List.Contains(Table.Column(Exceptions,toColumn),Record.Field(_,Column)) else List.Contains(Table.Column(Exceptions,toColumn),Record.Field(_,Column)))
in
	Selection

/*
let
	Source = #"Data (2)",
	Exceptions = tblExceptions,
	Column = "Account",
	ToColumn = Column,
	Filter = true,
	toColumn = if ToColumn = null then Column else ToColumn,
	Selection = Table.SelectRows(Source, each if Filter = null or Filter then not List.Contains(Table.Column(Exceptions,toColumn),Record.Field(_,Column)) else List.Contains(Table.Column(Exceptions,toColumn),Record.Field(_,Column)))
in
	Selection
*/
