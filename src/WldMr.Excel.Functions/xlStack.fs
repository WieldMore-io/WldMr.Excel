module WldMr.Excel.Range

open ExcelDna.Integration
open WldMr.Excel.Helpers


let getSize (x: obj[,]) =
  match x.GetLength 0, x.GetLength 1 with
  | 0, _ | _, 0 -> 0, 0
  | 1, 1 when ExcelMissing.Value :> obj = x.[0, 0] -> 0, 0
  | ls -> ls


[<ExcelFunction(Category= "WldMr.Range", Description= "Stack two arrays vertically")>]
let xlStackH (x:obj[,], y:obj[,]) =
  let x0, x1 = x |> getSize
  let y0, y1 = y |> getSize
  if x0 + y0 = 0 then
    ExcelError.ExcelErrorNA |> box
  else
    let res = Array2D.create (max x0 y0) (x1 + y1)  ("" |> box)
    if x0 > 0 then
      for i = 0 to x0 - 1 do
        for j = 0 to x1 - 1 do
          res.[i, j] <- x.[i,j]
    if y0 > 0 then
      for i = 0 to y0 - 1 do
        for j = 0 to y1 - 1 do
          res.[i, j + x1] <- y.[i,j]
    res |> box


[<ExcelFunction(Category= "WldMr.Range", Description= "Stack two arrays side by side")>]
let xlStackV (x:obj[,], y:obj[,]) =
  let x0, x1 = x |> getSize
  let y0, y1 = y |> getSize
  let res = Array2D.create (x0 + y0) (max x1 y1)  ("" |> box)
  if x0 + y0 = 0 then
    ExcelError.ExcelErrorNA |> box
  else
    if x0 > 0 then
      for i = 0 to x0 - 1 do
        for j = 0 to x1 - 1 do
          res.[i, j] <- x.[i,j]
    if y0 > 0 then
      for i = 0 to y0 - 1 do
        for j = 0 to y1 - 1 do
          res.[i + x0, j] <- y.[i,j]
    res |> box


[<ExcelFunction(Category= "WldMr.Range", Description= "Trim array")>]
let xlTrimEmpty (x:obj[,]) =
  let mutable last0 = -1
  let mutable last1 = -1
  let x0 = x.GetLength 0
  let x1 = x.GetLength 1

  while last0 = - 1 do
    for i = x0 - 1 downto 0 do
      for j = x1 - 1 downto 0 do
        match x.[i,j] with
        | ExcelEmpty _
        | ExcelString "" -> ()
        | _ ->
          last0 <- max i last0

  while last1 = - 1 do
    for j = x1 - 1 downto 0 do
      for i = last0 downto 0 do
        match x.[i,j] with
        | ExcelEmpty _
        | ExcelString "" -> ()
        | _ ->
          last1 <- max j last1

  x.[0..last0, 0..last1]


[<ExcelFunction(Category= "WldMr.Range", Description= "Trim #NA from end of array")>]
let xlTrimNA (x:obj[,]) =
  let x0, x1 = x |> getSize

  if x0 = 0 || x1 = 0 then
    [[]] |> array2D
  else
    let mutable last0 = -1
    let mutable last1 = -1

    for i = x0 - 1 downto 0 do
      for j = x1 - 1 downto 0 do
        match x.[i,j] with
        | ExcelEmpty _
        | ExcelError ExcelError.ExcelErrorNA -> ()
        | _ ->
          last0 <- max i last0

    if last0 <> -1 then
      for j = x1 - 1 downto 0 do
        for i = last0 downto 0 do
          match x.[i,j] with
          | ExcelMissing _
          | ExcelError ExcelError.ExcelErrorNA -> ()
          | _ ->
            last1 <- max j last1

    if min last0 last1 = -1 then
      [[]] |> array2D
    else
      x.[0..last0, 0..last1]

