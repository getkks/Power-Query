(Source as table, Indices as list, ColumnNamesNew as list) =>
let
    ColumnNamesOld = List.Transform( Indices, each Table.ColumnNames(Source){_} ), 
    ZippedList = List.Zip( { ColumnNamesOld, ColumnNamesNew } ),
    #"Renamed Columns" = Table.RenameColumns( Source, ZippedList )
in
    #"Renamed Columns"