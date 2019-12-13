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
