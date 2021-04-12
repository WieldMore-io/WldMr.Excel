module WldMr.Excel.Slice

open WldMr.Excel.Utilities
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
    and! base1_er = toRow |> parseArg XlObj.toInt -1 (fun e -> $"arg 'ToRow': {e}" )
    and! base1_sc = fromColumn |> parseArg XlObj.toInt 1 (fun e -> $"arg 'FromColumn': {e}" )
    and! base1_ec = toColumn |> parseArg XlObj.toInt -1 (fun e -> $"arg 'ToColumn': {e}" )
    let sr = base1_sr - 1
    let sc = base1_sc - 1
    let er = base1_er - 1
    let ec = base1_ec - 1
    let nRows = range.GetLength 0
    let nCols = range.GetLength 1
    let startRow = if sr >= 0 then sr else nRows + sr + 1
    let startCol = if sc >= 0 then sc else nCols + sc + 1
    let endRow = if er >= 0 then er else nRows + er + 1
    let endCol = if ec >= 0 then ec else nCols + ec + 1
    let slice = range.[startRow..endRow, startCol..endCol]
    if slice.LongLength = 0L then
      return ExcelError.ExcelErrorNA |> box
    else
      return slice |> box
  }
  res |> XlObj.ofValidation


