(
	Source as text,
	Search as text,
	Replace as text) as text =>
		Text.Combine(
			List.ReplaceMatchingItems(
				Text.ToList(Source),
				List.Zip({
					Text.ToList(Search),
					Text.ToList(Replace)})))