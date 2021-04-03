module WldMr.Excel.Range

open ExcelDna.Integration
open WldMr.Excel.Helpers


let getSize (x: obj[,]) =
  match x.GetLength 0, x.GetLength 1 with
  | 0, _ | _, 0 -> 0, 0
  | 1, 1 when x.[0, 0] = objMissing -> 0, 0
  | ls -> ls


[<ExcelFunction(Category= "WldMr.Range", Description= "Stack two arrays vertically")>]
let xlStackH (x:obj[,], y:obj[,]) =
  let x0, x1 = x |> getSize
  let y0, y1 = y |> getSize
  Array2D.init (max x0 y0) (x1 + y1)
    (fun i j ->
      if j < x1 then
        if i < x0 then x.[i, j] else objNA
      else
        if i < y0 then y.[i, j - x1] else objNA
    )


[<ExcelFunction(Category= "WldMr.Range", Description= "Stack two arrays vertically")>]
let xlStackV (x:obj[,], y:obj[,]) =
  let x0, x1 = x |> getSize
  let y0, y1 = y |> getSize
  Array2D.init (x0 + y0) (max x1 y1)
    (fun i j ->
      if i < x0 then
        if j < x1 then x.[i, j] else objNA
      else
        if j < y1 then y.[i - x0, j] else objNA
    )


let trimPredicate pred (x:obj[,]) =
  let x0, x1 = x |> getSize

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


[<ExcelFunction(Category= "WldMr.Range", Description= "Trim #NA cells from the end of array")>]
let xlTrimNA (x:obj[,]) =
  trimPredicate ((<>) objNA) x


[<ExcelFunction(Category= "WldMr.Range", Description= "Trim empty cells from the end of array")>]
let xlTrimEmpty (x:obj[,]) =
  trimPredicate (function | ExcelEmpty _ | ExcelString "" -> false | _ -> true) x
