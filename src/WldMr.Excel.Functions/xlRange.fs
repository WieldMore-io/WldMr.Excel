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
let xlRangeAnd (range1:xlObj[,], range2: xlObj[,], range3: xlObj[,], range4: xlObj[,])
  : xlObj[,]
  =
  let ranges = [range1; range2; range3; range4]
  xlRangeCommon "xlRangeAnd" (&&) true ranges


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise boolean OR for arrays")>]
let xlRangeOr(range1:xlObj[,], range2: xlObj[,], range3: xlObj[,], range4: xlObj[,])
  : xlObj[,]
  =
  let ranges = [range1; range2; range3; range4]
  xlRangeCommon "xlRangeOr" (||) false ranges


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise Max for arrays")>]
let xlRangeMax(
  number1: xlObj[,],
  number2: xlObj[,],
  number3: xlObj[,],
  number4: xlObj[,],
  number5: xlObj[,],
  number6: xlObj[,],
  number7: xlObj[,],
  number8: xlObj[,]
  )
  : xlObj[,]
  =

  let scalarF f1 f2 f3 f4 f5 f6 f7 f8 =
    let fs =
      [f1; f2; f3; f4; f5; f6; f7; f8]
      |> List.collect Option.toList
    match fs with
    | [] -> Error "no numbers"
    | _ -> fs |> List.reduce max |> XlObj.ofFloat |> Ok


  ArrayFunctionBuilder
    .Add("Number1", XlObj.toFloat |> XlObjParser.asOption, number1)
    .Add("Number2", XlObj.toFloat |> XlObjParser.asOption, number2)
    .Add("Number3", XlObj.toFloat |> XlObjParser.asOption, number3)
    .Add("Number4", XlObj.toFloat |> XlObjParser.asOption, number4)
    .Add("Number5", XlObj.toFloat |> XlObjParser.asOption, number5)
    .Add("Number6", XlObj.toFloat |> XlObjParser.asOption, number6)
    .Add("Number7", XlObj.toFloat |> XlObjParser.asOption, number7)
    .Add("Number8", XlObj.toFloat |> XlObjParser.asOption, number8)
    .EvalFunction scalarF
  |> FunctionCall.eval


[<ExcelFunction(Category= "WldMr Array", Description= "Element-wise Min for arrays")>]
let xlRangeMin(
  number1: xlObj[,],
  number2: xlObj[,],
  number3: xlObj[,],
  number4: xlObj[,],
  number5: xlObj[,],
  number6: xlObj[,],
  number7: xlObj[,],
  number8: xlObj[,]
  )
  : xlObj[,]
  =

  let scalarF f1 f2 f3 f4 f5 f6 f7 f8 =
    let fs =
      [f1; f2; f3; f4; f5; f6; f7; f8]
      |> List.collect Option.toList
    match fs with
    | [] -> Error "no numbers"
    | _ -> fs |> List.reduce min |> XlObj.ofFloat |> Ok


  ArrayFunctionBuilder
    .Add("Number1", XlObj.toFloat |> XlObjParser.asOption, number1)
    .Add("Number2", XlObj.toFloat |> XlObjParser.asOption, number2)
    .Add("Number3", XlObj.toFloat |> XlObjParser.asOption, number3)
    .Add("Number4", XlObj.toFloat |> XlObjParser.asOption, number4)
    .Add("Number5", XlObj.toFloat |> XlObjParser.asOption, number5)
    .Add("Number6", XlObj.toFloat |> XlObjParser.asOption, number6)
    .Add("Number7", XlObj.toFloat |> XlObjParser.asOption, number7)
    .Add("Number8", XlObj.toFloat |> XlObjParser.asOption, number8)
    .EvalFunction scalarF
  |> FunctionCall.eval
