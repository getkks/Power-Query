let
    Source = F[FolderToTable](P[Data Path]),
    #"Sheet to Table" = Table.TransformColumns(Source,{{"WorkSheet", F[SheetToTable], type table}}),
    #"Expanded WorkSheet" = F[ExpandColumn](#"Sheet to Table", "WorkSheet")
in
    #"Expanded WorkSheet"