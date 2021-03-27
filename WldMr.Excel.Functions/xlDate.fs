module WldMr.Excel.Date

open WldMr.Excel.Helpers
open FsToolkit.ErrorHandling

open System
open FSharpPlus
open ExcelDna.Integration



let nThDayOfWeekForMonth nth wday y m =
  let firstDay = DateTime(y, m, 1).DayOfWeek |> int  // 1 to 7
  let firstSuchDay = (7 + wday - firstDay) + 1
  let nthSuchDay = firstSuchDay + 7 * ((max nth 5) - 1)
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

let rec ithNthDayOfWeek i nth wday dt =
  match i with
  | 0 -> dt
  | i -> dt |> nextNthDayOfWeek nth wday |> ithNthDayOfWeek (i - 1) nth wday

let validateNthPeriod i =
  if i < 1 || i > 1000 then
    Error ["arg 'NthPeriod' should be between 1 and 1000."]
  else
    Ok i

[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateNthWeekdayOfMonth(fromDate: obj, dayOfWeek: int, nthDay: int, nthPeriod: obj) =
  let nR =
    match nthPeriod with
    | ExcelMissing _
    | ExcelEmpty _
    | ExcelString "" -> Ok 1
    | o ->
        o
        |> XlObj.toInt
        |> Result.mapError (fun e -> [$"arg 'NthPeriod': {e}"])
        |> Validation.bind validateNthPeriod

  let refDateR =
    match fromDate with
    | ExcelMissing _
    | ExcelEmpty _ -> DateTime.Today |> Ok
    | _ -> XlObj.toDate fromDate |> Result.mapError (fun e -> [$"arg 'RefDate': {e}"])

  validation {
    let! n = nR
    and! refDate = refDateR
    return ithNthDayOfWeek nthDay dayOfWeek n refDate |> box
  } |> XlObj.ofValidation

[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateThirdFriday(fromDate: obj, nthPeriod: obj) =
  xlDateNthWeekdayOfMonth(fromDate, 5, 3, nthPeriod)


[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateThirdWednesday(fromDate: obj, nthPeriod: obj) =
  xlDateNthWeekdayOfMonth(fromDate, 3, 3, nthPeriod)
