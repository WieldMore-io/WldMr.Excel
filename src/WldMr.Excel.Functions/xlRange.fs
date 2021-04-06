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
    "Selects a subrange of an array.\r\n" +
    "The arguments use the following convention:\r\n" +
    "  1   means the 1st row/column\r\n" +
    "       ... \r\n" +
    " -2   means the second-to-last row/column\r\n" +
    " -1   means the last row/column"
)>]
let xlSlice
  (
    [<ExcelArgument(Description="input range")>]
      range:obj[,],
    [<ExcelArgument(Description="the first row to return, defaults to 1")>]
      fromRow: obj,
    [<ExcelArgument(Description="the last row to return, defaults to -1")>]
      toRow: obj,
    [<ExcelArgument(Description="the first column to return, defaults to 1")>]
      fromColumn: obj,
    [<ExcelArgument(Description="the last column to return, defaults to -1")>]
      toColumn: obj
  ) =
  let res = validation {
    let! base1_sr = fromRow |> parseArg XlObj.toInt 1 (fun e -> $"arg 'FromRow': {e}" )
    and! er = toRow |> parseArg XlObj.toInt -1 (fun e -> $"arg 'ToRow': {e}" )
    and! base1_sc = fromColumn |> parseArg XlObj.toInt 1 (fun e -> $"arg 'FromColumn': {e}" )
    and! ec = toColumn |> parseArg XlObj.toInt -1 (fun e -> $"arg 'ToColumn': {e}" )
    let sr = base1_sr - 1
    let sc = base1_sc - 1
    let nRows = range.GetLength 0
    let nCols = range.GetLength 1
    let startRow = if sr >= 0 then sr else nRows + sr
    let startCol = if sc >= 0 then sc else nCols + sc
    let endRow = if er >= 0 then er - 1 else nRows + er
    let endCol = if ec >= 0 then ec - 1 else nCols + ec
    let slice = range.[startRow..endRow, startCol..endCol]
    if slice.LongLength = 0L then
      return ExcelError.ExcelErrorNA |> box
    else
      return slice |> box
  }
  res |> XlObj.ofValidation


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
