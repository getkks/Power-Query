(
	Arguments as list,
	Func as function,
	optional ResultFunc as function) =>
let
	Source = List.Transform(Arguments,each Function.Invoke(Func,_))
in
	if ResultFunc is null then Source else ResultFunc(Source)
