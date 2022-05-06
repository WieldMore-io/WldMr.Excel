module WldMr.Excel.Functions.Slice

open ExcelDna.Integration
open FsToolkit.ErrorHandling
open WldMr.Excel


[<ExcelFunction(Category= "WldMr Array",
  Description=
    "Selects a subrange of an array.\r\n" +
    "A variation on the built-in 'Offset' function.\r\n" +
    "The arguments use the following convention:\r\n" +
    "  1   means the 1st row/column\r\n" +
    "       ... \r\n" +
    " -2   means the second-to-last row/column\r\n" +
    " -1   means the last row/column"
)>]
let xlSlice
  (
    [<ExcelArgument(Description="input range")>]
      range:xlObj[,],
    [<ExcelArgument(Description="the first row to return, defaults to 1")>]
      fromRow: xlObj,
    [<ExcelArgument(Description="the last row to return, defaults to -1")>]
      toRow: xlObj,
    [<ExcelArgument(Description="the first column to return, defaults to 1")>]
      fromColumn: xlObj,
    [<ExcelArgument(Description="the last column to return, defaults to -1")>]
      toColumn: xlObj,
    [<ExcelArgument(Description="row step, defaults to 1")>]
      rowStep: xlObj,
    [<ExcelArgument(Description="column step, defaults to 1")>]
      colStep: xlObj
  ): xlObj[,]
  =
  result {
    let nRows = range.GetLength 0
    let nCols = range.GetLength 1
    let! sr = fromRow |> (XlObj.toInt |> XlObjParser.withDefault 1 |> XlObjParser.withArgName "FromRow")
    and! sc = fromColumn |> (XlObj.toInt |> XlObjParser.withDefault 1 |> XlObjParser.withArgName "FromColumn")
    and! er = toRow |> (XlObj.toInt |> XlObjParser.withDefault -1 |> XlObjParser.withArgName "ToRow")
    and! ec = toColumn |> (XlObj.toInt |> XlObjParser.withDefault -1 |> XlObjParser.withArgName "ToColumn")
    let startRow = sr + if sr >= 0 then -1 else nRows
    let startCol = sc + if sc >= 0 then -1 else nCols
    let endRow = er + if er >= 0 then -1 else nRows
    let endCol = ec + if ec >= 0 then -1 else nCols

    let! rowStep_ = rowStep |> (XlObj.toInt |> XlObjParser.withDefault 1 |> XlObjParser.withArgName "RowStep")
    and! colStep_ = colStep |> (XlObj.toInt |> XlObjParser.withDefault 1 |> XlObjParser.withArgName "ColStep")

    let totalRows = (endRow - startRow) / rowStep_ + 1
    let totalCols = (endCol - startCol) / colStep_ + 1

    let res =
      Array2D.init totalRows totalCols (fun i j ->
        range.[startRow + i * rowStep_, startCol + j * colStep_]
      )

    if res.LongLength = 0L then
      return XlObj.Error.xlNA |> XlObjRange.ofCell
    else
      return res
  }
  |> XlObjRange.ofResult
