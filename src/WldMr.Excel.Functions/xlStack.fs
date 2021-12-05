module WldMr.Excel.Functions.Stack

open ExcelDna.Integration
open WldMr.Excel


[<ExcelFunction(Category= "WldMr Array", Description= "Stack two arrays vertically")>]
let xlStackH (x:objCell[,], y:objCell[,]) =
  let x0, x1 = x |> XlObjRange.getSize
  let y0, y1 = y |> XlObjRange.getSize
  Array2D.init (max x0 y0) (x1 + y1)
    (fun i j ->
      if j < x1 then
        if i < x0 then x.[i, j] else XlObj.Error.xlNA
      else
        if i < y0 then y.[i, j - x1] else XlObj.Error.xlNA
    )


[<ExcelFunction(Category= "WldMr Array", Description= "Stack two arrays vertically")>]
let xlStackV (x:objCell[,], y:objCell[,]) =
  let x0, x1 = x |> XlObjRange.getSize
  let y0, y1 = y |> XlObjRange.getSize
  Array2D.init (x0 + y0) (max x1 y1)
    (fun i j ->
      if i < x0 then
        if j < x1 then x.[i, j] else XlObj.Error.xlNA
      else
        if j < y1 then y.[i - x0, j] else XlObj.Error.xlNA
    )


let private trimPredicate pred (x:objCell[,]) =
  let x0, x1 = x |> XlObjRange.getSize

  let lastRow =
    seq { for i in x0-1 .. -1 .. 0 do yield x.[i, *] }
    |> Seq.tryFindIndex (Array.exists pred)
    |> Option.defaultValue x0
    |> (-) (x0 - 1)

  let lastCol =
    seq { for j in x1-1 .. -1 .. 0 do yield x.[0..lastRow, j] }
    |> Seq.tryFindIndex (Array.exists pred)
    |> Option.defaultValue x1
    |> (-) (x1 - 1)

  x.[0..lastRow, 0..lastCol]


[<ExcelFunction(Category= "WldMr Array", Description= "Trim #NA cells from the end of array")>]
let xlTrimNA (x:objCell[,]) =
  trimPredicate ((<>) XlObj.Error.xlNA) x


[<ExcelFunction(Category= "WldMr Array", Description= "Trim empty cells from the end of array")>]
let xlTrimEmpty (x:objCell[,]) =
  trimPredicate (function | ExcelEmpty _ | ExcelString "" -> false | _ -> true) x
