module WldMr.Excel.SubRange

// open WldMr.CommonDataLogic.Net
open WldMr.Excel.Helpers
open FsToolkit.ErrorHandling
open System
open FSharpPlus
open ExcelDna.Integration


let parseArg parseF t (errMap: string -> string) (o:obj)=
  match o with
  | ExcelEmpty _ -> t |> Ok
  | ExcelEmpty _ -> t |> Ok
  | ExcelString "" -> t |> Ok
  | o ->
      o
      |> parseF
      |> Result.mapError (errMap >> fun x -> [x])


[<ExcelFunction(Category= "WldMr.Range",
  Description=
    "Selects a subrange of rows from the input range.
The 'First' and 'End' arguments use the following convention:
\t 0   means the first row
\t 1   means the second row
\t    ... and so on ...
\t-2   means the second-to-last row
\t-1   means the last row"
)>]
let xlRangeSubRows
  (
    [<ExcelArgument(Description="the range to select")>] 
      range:obj[,],
    [<ExcelArgument(Description="the first row to return")>] 
      start: obj,
    [<ExcelArgument(Description="the last row to return", Name="end")>] 
      finish: obj
  ) =
  let res = validation {
    let! s0 = start |> parseArg XlObj.toInt 0 (fun e -> $"arg 'start': {e}" )
    and! e0 = finish |> parseArg XlObj.toInt -1 (fun e -> $"arg 'start': {e}" )
    //and! i = i |> parseArg XlObj.toInt 1 (fun e -> $"arg 'start': {e}" )
    let nRows = range.GetLength 0
    let s = if s0 >= 0 then s0 else nRows + s0
    let e = if e0 >= 0 then e0 else nRows + e0
    let rowList = range |> Array2D.flattenArray |> fun l -> l.[s..e]
    do! rowList |> Result.requireNotEmpty ["No row selected"]
    return rowList |> array2D
  }
  res |>> box |> XlObj.ofValidation


//[<ExcelFunction(Category= "WldMr.Range",
//  Name="Range.Subrows",
//  Description=
//    "Selects a subrange of rows from the input range.
//The 'First' and 'End' arguments use the following convention:
//\t 0   means the first row
//\t 1   means the second row
//\t    ... and so on ...
//\t-2   means the second-to-last row
//\t-1   means the last row"
//)>]
//let renamedXlRangeSubRows(a: obj[,],b: obj, c:obj) =
//  xlRangeSubRows(a, b, c)


[<ExcelFunction(Category= "WldMr.Range",
  Description=
    "Selects a subrange of columns from the input range.
The 'First' and 'End' arguments use the following convention:
\t 0   means the first column
\t 1   means the second column
\t    ... and so on ...
\t-2   means the second-to-last column
\t-1   means the last column"
)>]
let xlRangeSubColumns
  (
    [<ExcelArgument(Description="the range to select")>] 
      range:obj[,],
    [<ExcelArgument(Description="the first column to return")>] 
      start: obj,
    [<ExcelArgument(Description="the last column to return", Name="end")>] 
      finish: obj
  ) =
  let res = validation {
    let! s0 = start |> parseArg XlObj.toInt 0 (fun e -> $"arg 'start': {e}" )
    and! e0 = finish |> parseArg XlObj.toInt -1 (fun e -> $"arg 'start': {e}" )
    //and! i = i |> parseArg XlObj.toInt 1 (fun e -> $"arg 'start': {e}" )
    let nCols = range.GetLength 1
    let s = if s0 >= 0 then s0 else nCols + s0
    let e = if e0 >= 0 then e0 else nCols + e0
    let colList = range |> Array2D.flattenArray |> List.map (fun l -> l.[s..e])
    do! match colList with
        | [] -> ["No column selected"] |> Error
        | x::_ -> x |> Result.requireNotEmpty ["No column selected"]
    return colList |> array2D
  }
  res |>> box |> XlObj.ofValidation



let boolOptionFold f zero bos =
  let bs = bos |> List.choose id
  match bs with
  | [] -> false
  | _ -> bs |> List.fold f zero 


let internal_and rows cols (rs: (Result<bool option, string list>)[,] list) =
  Array2D.init rows cols
    (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.get) |> boolOptionFold (&&) true)


[<ExcelFunction(Category= "WldMr.Range", Description= "Select a subrange of rows")>]
let xlRangeAnd (range1:obj[,], range2: obj[,], range3: obj[,], range4: obj[,]) =
  let getSize (r: obj[,]) = r.GetLength 0, r.GetLength 1
  let res = validation {
    let! rng1 = match range1 with | ExcelMissingRange _ -> ["'Range1' is missing"] |> Error | r -> r |> Some |> Ok
    let size = getSize rng1.Value
    let rng2 = match range2 with | ExcelMissingRange _ -> None | r -> r |> Some
    let rng3 = match range3 with | ExcelMissingRange _ -> None | r -> r |> Some
    let rng4 = match range4 with | ExcelMissingRange _ -> None | r -> r |> Some
    do! rng2 |> Option.forall (getSize >> (=) size) |> Result.requireTrue ["range2 must have the same size as range1"]
    do! rng3 |> Option.forall (getSize >> (=) size) |> Result.requireTrue ["range3 must have the same size as range1"]
    do! rng4 |> Option.forall (getSize >> (=) size) |> Result.requireTrue ["range4 must have the same size as range1"]
    let rngs = [rng1; rng2; rng3; rng4] |> List.choose id
    let rs = rngs |> List.map (Array2D.map XlObj.toBoolOption)

    return internal_and (fst size) (snd size) rs

  }
  let res2 = res |>> (Array2D.map box)
  let res3 = res2 |>> box
  res3 |> XlObj.ofValidation

