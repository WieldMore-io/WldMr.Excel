module WldMr.Excel.Example.TestFunction

open ExcelDna.Integration
open WldMr.Excel


[<ExcelFunction>]
let dnatestRangeSize (range:xlObj[,]): xlObj[,] =
  [| range.GetLength 0; range.GetLength 1 |]
  |> Array.map XlObj.ofInt
  |> XlObjRange.Column.ofSeqWithEmpty XlObj.Error.xlNA

