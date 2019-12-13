(
	Source as table,
	Exceptions as table,
	Columns as any,
	optional ToColumns as any,
	optional Filter as logical) =>
let
//	  Error = if not (Value.Is(Columns, type list) or Value.Is(Columns, type text)) then error "Function FilterExceptions: => Columns should be either list of column names or a single column name (text)." else null,
	ColumnList = Value.Is(Columns, type list),
	JoinType = if Filter <> null and not Filter then JoinKind.Inner else JoinKind.LeftAnti,
	toColumns = if ToColumns = null then Columns else ToColumns,
//	  MergeException = if Error <> null then Error else Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	MergeException =  Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	#"Removed Columns" = Table.RemoveColumns(MergeException,{"tblExceptions"})
in
	#"Removed Columns"

/*let
	Source = Data3,
	Exceptions = tblExceptions,
	Columns = "Account",
	ToColumns = "Account",
	Filter = false,
	Error = if not (Value.Is(Columns, type list) or Value.Is(Columns, type text)) then error "Function FilterExceptions: => Columns should be either list of column names or a single column name (text)." else null,
	ColumnList = Value.Is(Columns, type list),
	JoinType = if Filter <> null and not Filter then JoinKind.Inner else JoinKind.LeftAnti,
	toColumns = if ToColumns = null then Columns else ToColumns,
	MergeException = if Error <> null then Error else Table.NestedJoin(Source,if ColumnList then Columns else {Columns},Exceptions,if ColumnList then toColumns else {toColumns},"tblExceptions", JoinType),
	#"Removed Columns" = Table.RemoveColumns(MergeException,{"tblExceptions"})
in
	#"Removed Columns"
*/
