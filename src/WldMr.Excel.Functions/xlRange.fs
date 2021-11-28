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


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise boolean AND for arrays")>]
let xlRangeAnd (range1:objCell[,], range2: objCell[,], range3: objCell[,], range4: objCell[,]) =
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
      |> List.map (Array2D.map ( XlTypes.tag >> XlObj.toBoolOption))
      |> internalAnd (fst size) (snd size)
      |> Array2D.map XlObj.ofBool    // box the bools
  }
  res |> XlObjRange.ofValidation


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise boolean OR for arrays")>]
let xlRangeOr (range1:objCell[,], range2: objCell[,], range3: objCell[,], range4: objCell[,]) =
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
      |> List.map (Array2D.map ( XlTypes.tag >> XlObj.toBoolOption))
      |> internalOr (fst size) (snd size)
      |> Array2D.map XlObj.ofBool    // box the bools
  }
  res |> XlObjRange.ofValidation
