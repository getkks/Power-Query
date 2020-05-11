/*

Copyright 2019 - 2020

Project			: Power-Query		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query

File			: TransformList.pq		@ => d:\Karthik K Selvan\Power Query\Library Functions\Power-Query\TransformList.pq
Created			: Friday, 13th December 2019 4:58:19 pm

Author			: KarthikK Selvan

Change History:

*/
(
	Arguments as list,
	Func as function,
	optional ResultFunc as function) =>
let
	Source = List.Transform(Arguments,each Function.Invoke(Func,_))
in
	if ResultFunc is null then Source else ResultFunc(Source)
