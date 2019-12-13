(Source as table) as table =>
let
    #"Column Names" = Table.ColumnNames(Source),
    #"Trimmed and Cleaned Column Names" = List.Transform(#"Column Names",each Text.Clean(Text.Trim(_))),
    #"Renamed the Table" = Table.RenameColumns(Source,List.Zip({#"Column Names",#"Trimmed and Cleaned Column Names"}))
in
    #"Renamed the Table"
