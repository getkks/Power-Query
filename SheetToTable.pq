(Source as table,
	optional ColumnNumber as number,
	optional NoNullandBlank) =>

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

	Selected = Table.SelectRows(Source, each
		BlankandNull(Record.Field(_,
			"Column" & (
				if ColumnNumber = null then
					"4"
				else
					Number.ToText(ColumnNumber))),
			NoNullandBlank)),

	Promoted =	Table.PromoteHeaders(Selected, [PromoteAllScalars=true])
in
	Promoted