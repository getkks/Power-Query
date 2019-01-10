(Source as table, Column as text, Func as function, optional Type as type)=
let
    #Added Custom = Table.AddColumn(Source, Temp Column, each Func(_), Type),
    #ReOrder Columns = F[ReorderColumnsSubset](#Added Custom, {Temp Column}, List.PositionOf(Table.ColumnNames(#Added Custom),Column)),
    #Removed Column = Table.RemoveColumns(#ReOrder Columns,{Column}),
    #Renamed Columns = Table.RenameColumns(#Removed Column,{{Temp Column, Column}})
in
    #Renamed Columns