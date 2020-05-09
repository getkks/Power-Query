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

Project			: Power-Query @ d:\Development\M\Power-Query

File			: DateFromText.pq @ d:\Development\M\Power-Query\DateFromText.pq
File Created	: Saturday, 9th May 2020 3:20:38 pm

Author			: getkks (getkks@live.in)

Change History:

2020-05-09			getkks			Refactored the code
*/
(
	Source as text,
	optional Positions as list/*list meta [
		Documentation.FieldCaption = "Positions",
		Documentation.FieldDescription = "List of positions for date splitter.",
		Documentation.SampleValues =  {"{0,2,4}"}]*/,
	optional Combiner as function/*list meta [
		Documentation.FieldCaption = "Combiner",
		Documentation.FieldDescription = "Function which can combine day month and year into date as required.",
		Documentation.SampleValues = (li) => #date(li{2},li{0},li{1})]*/) as date=>
let
	Splitter = Splitter.SplitTextByPositions(
		if Positions is null then
			{0,2,4}
		else
			Positions),

	Split = List.Transform(Splitter(Source), Int16.From)
in
	if Combiner is null then
			#date(Split{2},Split{0},Split{1})
		else
			Combiner(Split)
