(Source) =>
let
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 0),
    #"Filtered Rows" = Table.SelectRows(#"Added Index", each List.Contains(SheetorTableName,[Index])),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Index"})
in
    #"Removed Columns"