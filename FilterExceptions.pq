(Source as table, Exceptions as table, Columns as any, optional ToColumns as any, optional Flip as logical) =>
let
    Flip = if Flip = null then false else true,
    ToColumns = if ToColumns = null then Columns else ToColumns,
    MergeException = if Value.Is(Columns, type list) then Table.NestedJoin(Source, Columns,Exceptions, ToColumns, "Exceptions", JoinKind.LeftOuter) else if Value.Is(Columns, type text) then Table.NestedJoin(Source, {Columns},Exceptions, {ToColumns}, "Exceptions", JoinKind.LeftOuter) else error "Function FilterExceptions: => Columns should be either list of column names or a single column name (text).",
    FilterException = Table.SelectRows(MergeException, each if Flip then Table.RowCount([Exceptions]) <> 0 else Table.RowCount([Exceptions]) = 0),
    RemoveException = Table.RemoveColumns(FilterException, {"Exceptions"})
in
    RemoveException