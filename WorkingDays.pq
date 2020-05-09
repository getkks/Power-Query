/*

Copyright 2020 - 2020 getkks, getkks@live.in

Unlicense

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http: //unlicense.org>

Project			: Power Query Library @ d:\Development\M\Power-Query

File			: WorkingDays.pq @ d:\Development\M\Power-Query\WorkingDays.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

*/
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
