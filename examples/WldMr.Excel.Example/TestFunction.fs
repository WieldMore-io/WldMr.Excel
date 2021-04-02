module WldMr.Excel.Example.TestFunction

open ExcelDna.Integration

[<ExcelFunction>]
let dnatestRangeSize (range:obj[,]): obj[,] =
  Array2D.init 1 2 (
    fun x y -> (if y = 0 then range.GetLength 0 else range.GetLength 1) |> box
  )
