let
	Shared = #shared,
	//cyclic reference...?
	//SharedPeriods = Record.Rename(Shared, (k,v) => Text.Replace(k, "_", ".")),

	SharedPeriods = Record.FromList(Record.FieldValues(Shared), List.Transform(Record.FieldNames(Shared), each Text.Replace(_, "_", "."))),

	SharedMerged = Record.Combine({Shared, SharedPeriods}),
	//If I can make a wrapper function to enable profiling/persistence, wrap these as well

	Files = Folder.Files(P[PQ Functions]),
	MFiles = Table.SelectRows(Files, each Text.Lower([Extension]) = ".pq"
		// and [Folder Path] = LoadPath    // non-recursive
	),
	NoExt = Table.TransformColumns(MFiles, {"Name", each Text.Start(_, Text.Length(_)-3)}),
	CustomNames = Table.Column(NoExt, "Name"),
	CustomBoth = List.Union({CustomNames, CustomNames}),
	Load = (fnName as text, optional BasePath as text) as function =>
		let
			DefaultPath = P[PQ Functions],

			BasePath = if (BasePath <> null) then BasePath else DefaultPath,
			Path = BasePath & (if Text.End(BasePath, 1) <> "\" then "\" else ""),
			File = Path & fnName & ".pq",

			Function = try Expression.Evaluate(Text.Replace(fnName, ".", "_"), #shared)  //if already imported into the workbook just use the existing one
				otherwise Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(File))), #shared)
		in
			Function,
	CustomLoaded = Record.FromList(List.Transform(CustomBoth, Load), CustomBoth),

	SharedAndLoaded = Record.Combine({CustomLoaded, SharedMerged})
	
in
	SharedAndLoaded