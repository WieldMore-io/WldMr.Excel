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


let thirdFridayForMonth y m =
  let firstDay = DateTime(y, m, 1).DayOfWeek |> int  // Sunday is 0
  let firstFriday = (6 - firstDay + 6) % 7 + 1
  DateTime(y, m, firstFriday + 14)


let thirdFriday (dt: DateTime) =
  let m = (12 - dt.Month) % 3 |> dt.AddMonths
  thirdFridayForMonth m.Year m.Month

let nextThirdFriday dt =
  let a = thirdFriday dt
  if a <= dt then
    dt.AddMonths(3) |> thirdFriday
  else
    a

let rec nthThirdFriday n dt =
  match n with
  | 0 -> dt
  | n -> dt |> nextThirdFriday |> nthThirdFriday (n - 1)


let validateNthPeriod i =
  if i < 1 || i > 1000 then
    Error ["arg 'NthPeriod' should be between 1 and 1000."]
  else
    Ok i


[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateThirdFriday(fromDate: obj, nthPeriod: obj) =
  let nR =
    match nthPeriod with
    | ExcelMissing _
    | ExcelEmpty _
    | ExcelString "" -> Ok 1
    | o ->
        o
        |> XlObj.toInt
        |> Result.mapError (sprintf "arg 'NthPeriod': %s" >> fun x -> [x])
        |> Validation.bind validateNthPeriod

  let refDateR =
    match fromDate with
    | ExcelMissing _
    | ExcelEmpty _ -> DateTime.Today |> Ok
    | _ -> XlObj.toDate fromDate |> Result.mapError (sprintf "arg 'RefDate': %s" >> (fun x -> [x]))

  validation {
    let! n = nR
    and! refDate = refDateR
    return nthThirdFriday n refDate |> box
  } |> XlObj.ofValidation

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
        |> Result.mapError (sprintf "arg 'NthPeriod': %s" >> fun x -> [x])
        |> Validation.bind validateNthPeriod

  let refDateR =
    match fromDate with
    | ExcelMissing _
    | ExcelEmpty _ -> DateTime.Today |> Ok
    | _ -> XlObj.toDate fromDate |> Result.mapError (sprintf "arg 'RefDate': %s" >> (fun x -> [x]))

  validation {
    let! n = nR
    and! refDate = refDateR
    return ithNthDayOfWeek nthDay dayOfWeek n refDate |> box
  } |> XlObj.ofValidation
