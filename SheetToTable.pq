(Source as table,
	optional ColumnNumber as number,
	optional NoNullandBlank,
	optional SkipRows) =>

let
	BlankandNull = (Value,
		optional NullandBlank) =>

		let
			NullandBlank =
				if NullandBlank = null then
					true
				else
					NullandBlank,

			Result =
				if NullandBlank then
					Value <> null and Value <> ""
				else
					Value <> null
		in
			Result,

	Selected = if SkipRows = null then
			Table.SelectRows(Source, each
				BlankandNull(Record.Field(_, "Column" & (
					if ColumnNumber = null then
						"4"
					else
						Number.ToText(ColumnNumber)))
					,NoNullandBlank))
		else
			Table.Skip(Source, SkipRows),

	Promoted = Table.PromoteHeaders(Selected, [PromoteAllScalars=true])
in
	Promoted
