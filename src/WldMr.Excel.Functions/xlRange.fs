module WldMr.Excel.Functions.Range

open ExcelDna.Integration
open FsToolkit.ErrorHandling
open WldMr.Excel


let boolOptionFold f zero bos =
  let bs = bos |> List.choose id
  match bs with
  | [] -> false
  | _ -> bs |> List.fold f zero


let xlRangeCommon fnName booleanOp opZero ranges =

  let internalOp rows cols (rs: Result<bool option, string>[,] list) =
    Array2D.init rows cols
      (fun i j ->
        rs
        |> List.traverseResultM (fun r -> r.[i, j])
        |> Result.map ( boolOptionFold booleanOp opZero) )

  let optionRange = function ExcelMissingRange _ -> None | r -> r |> Some
  let ranges_ = ranges |> List.map optionRange |> List.choose id

  result {
    let! headRng =
      ranges_
      |> List.tryHead
      |> Option.fold (fun _-> Ok) (Error $"{fnName} needs at least one parameter")
    let size = XlObjRange.getSize headRng
    do! ranges_ |> List.forall (XlObjRange.getSize >> (=) size) |> Result.requireTrue "All ranges must have the same size"
    return
      ranges_
      |> List.map (Array2D.map XlObj.toBoolOption)
      |> internalOp (fst size) (snd size)
      |> Array2D.map ( Result.map XlObj.ofBool >> XlObj.ofResult)
  } |> XlObjRange.ofResult


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise boolean AND for arrays")>]
let xlRangeAnd (range1:objCell[,], range2: objCell[,], range3: objCell[,], range4: objCell[,])
  : objCell[,]
  =
  let ranges = [range1; range2; range3; range4]
  xlRangeCommon "xlRangeAnd" (&&) true ranges


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise boolean OR for arrays")>]
let xlRangeOr(range1:objCell[,], range2: objCell[,], range3: objCell[,], range4: objCell[,])
  : objCell[,]
  =
  let ranges = [range1; range2; range3; range4]
  xlRangeCommon "xlRangeOr" (||) false ranges

