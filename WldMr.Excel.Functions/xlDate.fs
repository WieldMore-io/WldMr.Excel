module WldMr.Excel.Date

open WldMr.Excel.Helpers
open FsToolkit.ErrorHandling

open System
open FSharpPlus
open ExcelDna.Integration


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
    Error "nthPeriod should be between 1 and 1000"
  else
    Ok i

[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateThirdFriday(fromDate: obj, nthPeriod: obj) =
  let nR = 
    match nthPeriod with
    | ExcelMissing _ -> Ok 1
    | ExcelEmpty _ -> Ok 1
    | ExcelString "" -> Ok 1
    | o -> o |> XlObj.toInt |> Result.mapError (sprintf "arg 'NthPeriod': %s") |> Result.bind validateNthPeriod
   
  let refDateR = 
    match fromDate with
    | ExcelMissing m -> DateTime.Today |> Result.Ok
    | ExcelEmpty _ -> DateTime.Today |> Result.Ok
    | _ -> XlObj.toDate fromDate |> Result.mapError (sprintf "arg 'RefDate': %s")

  let dateRes = monad {
    let! n = nR
    let! refDate = refDateR
    nthThirdFriday n refDate
  }
  dateRes |>> box |> XlObj.ofResult



[<ExcelFunction(Category= "WldMr.Date", Description= "Returns the next quarterly third friday")>]
let xlDateThirdFridayNew(fromDate: obj, nthPeriod: obj) =
  let nR = 
    match nthPeriod with
    | ExcelMissing _ -> Ok 1
    | ExcelEmpty _ -> Ok 1
    | ExcelString "" -> Ok 1
    | o ->
        o
        |> XlObj.toInt
        |> Result.mapError (sprintf "arg 'NthPeriod': %s" >> fun x -> [x])
        |> Result.bind (validateNthPeriod >> Result.mapError (fun x -> [x]))
   
  let refDateR = 
    match fromDate with
    | ExcelMissing m -> DateTime.Today |> Result.Ok
    | ExcelEmpty _ -> DateTime.Today |> Result.Ok
    | _ -> XlObj.toDate fromDate |> Result.mapError (sprintf "arg 'RefDate': %s" >> (fun x -> [x]))

  let dateRes = validation {
    let! n = nR
    and! refDate = refDateR
    return nthThirdFriday n refDate
  }
  dateRes |>> box |> XlObj.ofValidation
