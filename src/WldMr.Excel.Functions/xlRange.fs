module WldMr.Excel.SubRange

open WldMr.Excel.Helpers
open FsToolkit.ErrorHandling
open FSharpPlus
open ExcelDna.Integration


let parseArg parseF t (errMap: string -> string) (o:obj)=
  match o with
  | ExcelMissing _
  | ExcelEmpty _
  | ExcelString "" -> t |> Ok
  | o ->
      o
      |> parseF
      |> Result.mapError (errMap >> fun x -> [x])


[<ExcelFunction(Category= "WldMr.Range",
  Description=
    "Selects a subrange of rows from the input range.\r\n" +
      "The 'Start' and 'End' arguments use the following convention:\r\n" +
      "  0   means the first row\r\n" +
      "  1   means the second row\r\n" +
      "       ... and so on ...\r\n" +
      "  t-2   means the second-to-last row\r\n" +
      "  t-1   means the last row"
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
    let! s0 = start |> parseArg XlObj.toInt 0 (fun e -> $"arg 'Start': {e}" )
    and! e0 = finish |> parseArg XlObj.toInt -1 (fun e -> $"arg 'End': {e}" )
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
The 'Start' and 'End' arguments use the following convention:
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
    let! s0 = start |> parseArg XlObj.toInt 0 (fun e -> $"arg 'Start': {e}" )
    and! e0 = finish |> parseArg XlObj.toInt -1 (fun e -> $"arg 'End': {e}" )
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


[<ExcelFunction(Category= "WldMr.Range", Description= "Element-wise boolean AND for arrays")>]
let xlRangeAnd (range1:obj[,], range2: obj[,], range3: obj[,], range4: obj[,]) =
  let internalAnd rows cols (rs: (Result<bool option, string list>)[,] list) =
    Array2D.init rows cols
      (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.get) |> boolOptionFold (&&) true)
  let getSize (r: obj[,]) = r.GetLength 0, r.GetLength 1
  let optionRange = function ExcelMissingRange _ -> None | r -> r |> Some
  let rngs = [range1; range2; range3; range4] |>> optionRange |> List.choose id
  let res = validation {
    let! headRng = rngs |> List.tryHead |> Option.toResult |> Result.withError ["xlRangeOr needs at least one parameter"]
    let size = getSize headRng
    do! rngs |> List.forall (getSize >> (=) size) |> Result.requireTrue ["All ranges must have the same size"]
    return
      rngs
      |> List.map (Array2D.map XlObj.toBoolOption)
      |> internalAnd (fst size) (snd size)
      |> Array2D.map box    // box the bools
      |> box                // box the array
  }
  res |> XlObj.ofValidation


[<ExcelFunction(Category= "WldMr.Range", Description= "Element-wise boolean OR for arrays")>]
let xlRangeOr (range1:obj[,], range2: obj[,], range3: obj[,], range4: obj[,]) =
  let internalOr rows cols (rs: (Result<bool option, string list>)[,] list) =
    Array2D.init rows cols
      (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.get) |> boolOptionFold (||) false)
  let getSize (r: obj[,]) = r.GetLength 0, r.GetLength 1
  let optionRange = function ExcelMissingRange _ -> None | r -> r |> Some
  let rngs = [range1; range2; range3; range4] |>> optionRange |> List.choose id
  let res = validation {
    let! headRng = rngs |> List.tryHead |> Option.toResult |> Result.withError ["xlRangeOr needs at least one parameter"]
    let size = getSize headRng
    do! rngs |> List.forall (getSize >> (=) size) |> Result.requireTrue ["All ranges must have the same size"]
    return
      rngs
      |> List.map (Array2D.map XlObj.toBoolOption)
      |> internalOr (fst size) (snd size)
      |> Array2D.map box    // box the bools
      |> box                // box the array
  }
  res |> XlObj.ofValidation
