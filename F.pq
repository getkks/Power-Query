let
	Files = Folder.Files(
		P[PQ Functions]),

	#"Filtered Rows" = Table.SelectRows(Files, each Text.Lower([Extension]) = ".pq"),

	#"Added Query" = Table.AddColumn(#"Filtered Rows", "Query", each
		Text.Start([Name], Text.Length([Name])-Text.Length([Extension])),
		type text),

	#"Added Function" = Table.AddColumn(#"Added Query", "Function", each
		Expression.Evaluate(
			Text.FromBinary(
				File.Contents(
					[Folder Path] & [Name])),
		#shared)),

	Buffer = Table.Buffer(#"Added Function"),

	#"To Record" = Record.FromList(
		Buffer[Function],
		Buffer[Query])
in
	#"To Record"