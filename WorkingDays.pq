(start as date,
	end as date,
	optional HolidaysList as list) as number =>

let
	LstOfHolidays = if HolidaysList = null then {} else HolidaysList,
	NumLstOfHolidays = List.Transform(LstOfHolidays ,each Number.From(_)),
	LstOfDays = if start<end then {Number.From(start)..Number.From(end)} else {Number.From(end)..Number.From(start)},
	LstDiff  = List.Difference(LstOfDays, NumLstOfHolidays),
	LstMod = List.Transform(LstDiff, each Number.Mod(_, 7)),
	LstSel = List.Select(LstMod, each _>1) ,
	Result = if start<end then List.Count(LstSel)-1 else if start=end then 0 else List.Count(LstSel)*(-1)+1
in
	Result
