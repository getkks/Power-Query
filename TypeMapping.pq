() =>
	#table(
		type table
		[
			#"Type Name" = text, 
			Type = type
		],
		{
			{ "Text", type text},
			{ "String", type text},
			{ "Number", type number},
			{ "Decimal", type number},
			{ "Decimal Number", type number},
			{ "Currency", Currency.Type},
			{ "Int64", Int64.Type},
			{ "Whole Number", Int64.Type},
			{ "Percentage", Percentage.Type},
			{ "DateTime", type datetime},
			{ "Date", type date},
			{ "Time", type time},
			{ "DateTimeZone", type datetimezone},
			{ "Duration", type duration},
			{ "Logical", type logical},
			{ "Boolean", type logical},
			{ "Type", type type},
			{ "", type any},
			{ null, type any},
			{ "List", type list},
			{ "Record", type record}})