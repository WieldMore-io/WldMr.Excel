module WldMr.Excel.Functions.Stack

open ExcelDna.Integration
open WldMr.Excel

open FsToolkit.ErrorHandling

let private trimPredicate pred (x:xlObj[,]) =
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
let xlTrimNA (x:xlObj[,]) =
  trimPredicate ((<>) XlObj.Error.xlNA) x


[<ExcelFunction(Category= "WldMr Array", Description= "Trim empty cells from the end of array")>]
let xlTrimEmpty (x:xlObj[,]) =
  trimPredicate (function | ExcelEmpty _ | ExcelString "" -> false | _ -> true) x


[<RequireQualifiedAccess>]
type StackParameter =
  | Trim of string

let stackParameter (rng: xlObj[,]) =
  option {
    let! xlObjV =
      match XlObjRange.getSize rng with
      | 1, 1 -> Some rng.[0, 0]
      | _ -> None
    let! v = xlObjV |> XlObj.toString |> Result.fold Some (fun _ -> None)
    let! vl = v.ToLower()
    let! trimParam =
      if vl.StartsWith("trim:") then
        StackParameter.Trim vl.[5..] |> Some
      else
        None
    return trimParam
  }


let xlStackH_internal (parameter: StackParameter) (ranges: xlObj[,] list): xlObj[,] =
  let trimFunction =
    option {
      let (StackParameter.Trim trimStr) = parameter
      return
        match trimStr.ToLower() with
        | "none" -> id
        | "empty" -> xlTrimEmpty
        | "na" -> trimPredicate (function | ExcelEmpty _ | ExcelString "" -> false | x -> x <> XlObj.Error.xlNA)
        | _ -> id
    } |> Option.defaultValue id
  let processedRanges =
    ranges |> List.map trimFunction

  let nRanges = processedRanges |> List.length
  let sizes = processedRanges |> List.map XlObjRange.getSize
  let cols = sizes |> List.sumBy snd
  let colStarts =
    sizes
    |> List.truncate (nRanges-1)
    |> List.map snd
    |> List.scan (+) 0
  let rows = sizes |> List.map fst |> List.max

  let r = Array2D.create rows cols XlObj.Error.xlNA

  (colStarts, processedRanges)
  ||> List.iter2 (fun colStart rng ->
      let x, y = rng |> XlObjRange.getSize
      for i = 0 to x-1 do
        for j = 0 to y-1 do
          r.[i, colStart+j] <- rng.[i, j]
    )
  r


[<ExcelFunction(Category= "WldMr Array", Description= "Stack ranges horizontally")>]
let xlStackH (
  rng1: xlObj[,],
  rng2: xlObj[,],
  rng3: xlObj[,],
  rng4: xlObj[,],
  rng5: xlObj[,],
  rng6: xlObj[,],
  rng7: xlObj[,],
  rng8: xlObj[,],
  rng9: xlObj[,],
  rng10: xlObj[,],
  rng11: xlObj[,],
  rng12: xlObj[,],
  rng13: xlObj[,],
  rng14: xlObj[,],
  rng15: xlObj[,],
  rng16: xlObj[,],
  rng17: xlObj[,],
  rng18: xlObj[,],
  rng19: xlObj[,]
  ) =
  let ranges = [
    rng1; rng2; rng3; rng4; rng5; rng6; rng7; rng8; rng9;
    rng10; rng11; rng12; rng13; rng14; rng15; rng16; rng17; rng18; rng19;
  ]
  let revRanges = ranges |> List.rev |> List.skipWhile (fun rng -> rng |> XlObjRange.isMissing)

  match revRanges with
  | [] ->
      XlObj.Error.xlNA |> XlObjRange.ofCell
  | lastRange::rest ->
      match lastRange |> stackParameter with
      | None ->
          xlStackH_internal (StackParameter.Trim "empty" ) (List.rev revRanges)
      | Some parameter ->
          xlStackH_internal parameter (List.rev rest)


let xlStackV_internal (parameter: StackParameter) (ranges: xlObj[,] list): xlObj[,] =
  let trimFunction =
    option {
      let (StackParameter.Trim trimStr) = parameter
      return
        match trimStr.ToLower() with
        | "none" -> id
        | "empty" -> xlTrimEmpty
        | "na" -> trimPredicate (function | ExcelEmpty _ | ExcelString "" -> false | x -> x <> XlObj.Error.xlNA)
        | _ -> id
    } |> Option.defaultValue id
  let processedRanges =
    ranges |> List.map trimFunction

  let nRanges = processedRanges |> List.length
  let sizes = processedRanges |> List.map XlObjRange.getSize
  let rows = sizes |> List.sumBy fst
  let rowStarts =
    sizes
    |> List.truncate (nRanges-1)
    |> List.map fst
    |> List.scan (+) 0
  let cols = sizes |> List.map snd |> List.max

  let r = Array2D.create rows cols XlObj.Error.xlNA

  (rowStarts, processedRanges)
  ||> List.iter2 (fun rowStart rng ->
      let x, y = rng |> XlObjRange.getSize
      for i = 0 to x-1 do
        for j = 0 to y-1 do
          r.[rowStart+i, j] <- rng.[i, j]
    )

  r


[<ExcelFunction(Category= "WldMr Array", Description= "Stack ranges vertically")>]
let xlStackV (
  rng1: xlObj[,],
  rng2: xlObj[,],
  rng3: xlObj[,],
  rng4: xlObj[,],
  rng5: xlObj[,],
  rng6: xlObj[,],
  rng7: xlObj[,],
  rng8: xlObj[,],
  rng9: xlObj[,],
  rng10: xlObj[,],
  rng11: xlObj[,],
  rng12: xlObj[,],
  rng13: xlObj[,],
  rng14: xlObj[,],
  rng15: xlObj[,],
  rng16: xlObj[,],
  rng17: xlObj[,],
  rng18: xlObj[,],
  rng19: xlObj[,]
  ) =
  let ranges = [
    rng1; rng2; rng3; rng4; rng5; rng6; rng7; rng8; rng9;
    rng10; rng11; rng12; rng13; rng14; rng15; rng16; rng17; rng18; rng19;
  ]
  let revRanges = ranges |> List.rev |> List.skipWhile (fun rng -> rng |> XlObjRange.isMissing)

  match revRanges with
  | [] ->
      XlObj.Error.xlNA |> XlObjRange.ofCell
  | lastRange::rest ->
      match lastRange |> stackParameter with
      | None ->
          xlStackV_internal (StackParameter.Trim "empty" ) (List.rev revRanges)
      | Some parameter ->
          xlStackV_internal parameter (List.rev rest)
