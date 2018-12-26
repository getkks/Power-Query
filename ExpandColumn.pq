(tblData as table, Column as text, optional KeepOtherColumns as logical) =>
    let
        KeepOtherColumns = if KeepOtherColumns = null then true else KeepOtherColumns,
        Expanded = if KeepOtherColumns then Table.ExpandTableColumn(tblData, Column, Table.ColumnNames(Table.Column(tblData,Column){0})) else Table.ExpandTableColumn(Table.SelectColumns(tblData,{Column}), Column, Table.ColumnNames(Table.Column(tblData,Column){0}))
    in
        Expanded