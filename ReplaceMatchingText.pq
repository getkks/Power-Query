/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: ReplaceMatchingText.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\ReplaceMatchingText.pq
Created			: Wednesday, 18th September 2019 4:25:11 pm

Author			: KarthikK Selvan

Change History:

*/
(
	Source as text,
	Search as text,
	Replace as text) as text =>
		if Source = null then
			null
		else
			Text.Combine(
				List.ReplaceMatchingItems(
					Text.ToList(Source),
					List.Zip({
						Text.ToList(Search),
						Text.ToList(Replace)})))
