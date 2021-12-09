module WldMr.Excel.Date

open WldMr.Excel
open FsToolkit.ErrorHandling

open System
open ExcelDna.Integration


module NthWeekdayOfMonth =
  let nThDayOfWeekForMonth nth wday y m =
    let firstDay = DateTime(y, m, 1).DayOfWeek |> int  // 1 to 7
    let firstSuchDay = (7 + wday - firstDay) % 7 + 1
    let nthSuchDay = firstSuchDay + 7 * ((min nth 5) - 1)
    if nthSuchDay > DateTime.DaysInMonth(y, m) then
      DateTime(y, m, nthSuchDay - 7)
    else
      DateTime(y, m, nthSuchDay)

  let nthDayOfWeek nth wday (dt: DateTime) =
    let m = (12 - dt.Month) % 3 |> dt.AddMonths
    nThDayOfWeekForMonth nth wday m.Year m.Month

  let nextNthDayOfWeek nth wday dt =
    let a = nthDayOfWeek nth wday dt
    if a <= dt then
      dt.AddMonths(3) |> nthDayOfWeek nth wday
    else
      a

  let rec ithNthDayOfWeek nth wday i dt =
    match i with
    | 0 -> dt
    | i -> dt |> nextNthDayOfWeek nth wday |> ithNthDayOfWeek nth wday (i - 1)


[<ExcelFunction(Category= "WldMr Date", Description= "Returns the next quarterly nth given day of the week")>]
let xlDateNthWeekdayOfMonth
  (
    [<ExcelArgument(Description="date to start from, today if missing")>]
      refDate: xlObj[,],
    [<ExcelArgument(Description="Monday: 1\r\nTuesday: 2\r\n...\r\nSunday: 7 (or 0)")>]
      dayOfWeek: xlObj,
    [<ExcelArgument(Description="1 to return the first given day\r\n...\r\n4 to return the fourth\r\n5 to return the last")>]
      nthSuchDay: xlObj,
    [<ExcelArgument(Description="1 to return the first quarterly such date following the reference date, 2 to return the second, ...")>]
      nthPeriod: xlObj[,]
  ): xlObj[,] =
  let dateNthWeekdayOfMonth refDate nthPeriod =
    result {
      let! dow = dayOfWeek |> (XlObj.toInt |> XlObj.withArgName "DayOfWeek")
      let! nth = nthSuchDay |> (XlObj.toInt |> XlObj.withArgName "NthSuchDay")

      do! (0 < nth && nth < 6) |> Result.requireTrue "arg 'NthSuchDay' should be between 1 and 5."
      do! (0 < nthPeriod && nthPeriod < 1001) |> Result.requireTrue "arg 'NthPeriod' should be between 1 and 1000."

      return NthWeekdayOfMonth.ithNthDayOfWeek nth (dow-1) nthPeriod refDate |> XlObj.ofDate
    }

  ArrayFunctionBuilder
    .Add("RefDate", XlObj.toDate |> XlObj.withDefault DateTime.Today, refDate)
    .Add("NthPeriod", XlObj.toInt |> XlObj.withDefault 1, nthPeriod)
    .EvalFunction dateNthWeekdayOfMonth
  |> FunctionCall.catchExceptions
  |> FunctionCall.eval


[<ExcelFunction(Category= "WldMr Date", Description= "Returns the next quarterly third Friday")>]
let xlDateThirdFriday(fromDate: xlObj[,], nthPeriod: xlObj[,]): xlObj[,] =
  xlDateNthWeekdayOfMonth(fromDate, 5.0 |> XlObj.ofFloat, 3.0 |> XlObj.ofFloat, nthPeriod)


[<ExcelFunction(Category= "WldMr Date", Description= "Returns the next quarterly third Wednesday")>]
let xlDateThirdWednesday(fromDate: xlObj[,], nthPeriod: xlObj[,]): xlObj[,] =
  xlDateNthWeekdayOfMonth(fromDate, 3.0 |> XlObj.ofFloat, 3.0 |> XlObj.ofFloat, nthPeriod)
