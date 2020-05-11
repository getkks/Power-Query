/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: SheetToTable.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\SheetToTable.pq
Created			: Tuesday, 30th April 2019 5:32:41 pm

Author			: KarthikK Selvan

Change History:

*/
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
			if SkipRows = 0 then
				if Comparer.Equals(Comparer.OrdinalIgnoreCase, Table.FirstValue(Source), "UTF8") then
					Table.Skip(Source, 1)
				else
					Source
			else
				Table.Skip(Source, SkipRows),

	Promoted = Table.PromoteHeaders(Selected, [PromoteAllScalars=true])
in
	Promoted
