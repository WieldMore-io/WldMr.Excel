module WldMr.Excel.SubRange

open WldMr
open WldMr.Excel.Utilities
open FsToolkit.ErrorHandling
open ExcelDna.Integration
open WldMr


let boolOptionFold f zero bos =
  let bs = bos |> List.choose id
  match bs with
  | [] -> false
  | _ -> bs |> List.fold f zero


let internal_and rows cols (rs: (Result<bool option, string list>)[,] list) =
  Array2D.init rows cols
    (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.value) |> boolOptionFold (&&) true)


[<ExcelFunction(Category= "WldMr.Range", Description= "Element-wise boolean AND for arrays")>]
let xlRangeAnd (range1:obj[,], range2: obj[,], range3: obj[,], range4: obj[,]) =
  let internalAnd rows cols (rs: (Result<bool option, string list>)[,] list) =
    Array2D.init rows cols
      (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.value) |> boolOptionFold (&&) true)
  let optionRange = function ExcelMissingRange _ -> None | r -> r |> Some
  let rngs = [range1; range2; range3; range4] |> List.map optionRange |> List.choose id
  let res = validation {
    let! headRng = rngs |> List.tryHead |> Option.toResult id |> Result.withError ["xlRangeOr needs at least one parameter"]
    let size = XlObj.getSize headRng
    do! rngs |> List.forall (XlObj.getSize >> (=) size) |> Result.requireTrue ["All ranges must have the same size"]
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
      (fun i j -> rs |> List.map (fun r -> r.[i, j] |> Result.value) |> boolOptionFold (||) false)
  let optionRange = function ExcelMissingRange _ -> None | r -> r |> Some
  let rngs = [range1; range2; range3; range4] |> List.map optionRange |> List.choose id
  let res = validation {
    let! headRng = rngs |> List.tryHead |> Option.toResult id |> Result.withError ["xlRangeOr needs at least one parameter"]
    let size = XlObj.getSize headRng
    do! rngs |> List.forall (XlObj.getSize >> (=) size) |> Result.requireTrue ["All ranges must have the same size"]
    return
      rngs
      |> List.map (Array2D.map XlObj.toBoolOption)
      |> internalOr (fst size) (snd size)
      |> Array2D.map box    // box the bools
      |> box                // box the array
  }
  res |> XlObj.ofValidation
