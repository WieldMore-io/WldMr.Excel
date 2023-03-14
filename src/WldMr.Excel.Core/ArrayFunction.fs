namespace WldMr.Excel

open FsToolkit.ErrorHandling
open WldMr.Excel


[<RequireQualifiedAccess>]
module ArrayFunction =

  type ArgValue<'T> =
    | NotCalculated
    | Scalar of 'T
    | Row of 'T[]
    | Col of 'T[]

  let private ofRowColInternal x =
    x
    |> Result.map (Option.fold (fun _ -> fst) 1)

  let inline private flip f a b = f b a

  let private validateDim dim (name: string) currentDim =
    match dim, currentDim with
    | 1, _ -> currentDim |> Ok
    | n, Some (p, _) when n = p -> currentDim |> Ok
    | _, Some (_, currentName) ->
        Error $"Argument '{name}' and '{currentName}' have inconsistent dimensions."
    | _, None -> (dim, name) |> Some |> Ok


  let private getPseudoElt (i: int) (j: int) a =
    let l1 = a |> Array2D.length1
    let l2 = a |> Array2D.length2
    a.[ (if l1 = 1 then 0 else i), (if l2 = 1 then 0 else j) ]


  type UdfArrayArgBase(name: string, value: xlObj[,], rest: UdfArrayArgBase option ) =
    member internal uaa.Name = name
    member internal uaa.Value = value
    member internal uaa.Rest = rest

    member internal uaa.RowsInternal(): Result<(int * string) option, string> =
      result {
        let! recDim = uaa.Rest |> Option.map (fun x -> x.RowsInternal ()) |> Option.defaultValue (Ok None)
        return! validateDim (uaa.Value |> Array2D.length1) uaa.Name recDim
      }
    member internal uaa.Rows () = uaa.RowsInternal() |> ofRowColInternal

    member internal uaa.ColsInternal(): Result<(int * string) option, string> =
      result {
        let! recDim = uaa.Rest |> Option.map (fun x -> x.ColsInternal ()) |> Option.defaultValue (Ok None)
        return! validateDim (uaa.Value |> Array2D.length2) uaa.Name recDim
      }
    member internal uaa.Cols () = uaa.ColsInternal() |> ofRowColInternal

    member internal uaa.returnArray2d f () =
      result {
        let! nRows = uaa.Rows()
        let! nCols = uaa.Cols()
        let a: xlObj[,] = Array2D.init nRows nCols (fun i j -> f i j |> Result.bind id |> XlObj.ofResult)
        return a |> XlObjRange.ofArray2dWithEmpty (XlObj.ofErrorMessage "empty.")
      } |> XlObjRange.ofResult


    member internal uaa.eval f subEval conversion i j =
      result {
        let! fp = subEval f i j
        let! v = uaa.Value |> getPseudoElt i j |> (XlObjParser.withArgName uaa.Name conversion)
        return fp v
      }

    member internal uaa.returnArray2dFromArrays f preferHorizontal ()  =
      result {
        let! nRows = uaa.Rows()
        let! nCols = uaa.Cols()
        do! (nRows = 1 || nCols = 1) |> Result.requireTrue "2-dimensional ranges are not accepted."

        let resLength = max nRows nCols
        let res =
          [|
             for i in 0 .. resLength - 1 ->
               let a = f (min nRows i) (min nCols i) |> Result.bind id
               match a with
               | Error e -> [| Error e |]
               | Ok [||] -> [| Error "Empty"|]
               | Ok a -> a |> Array.map Ok
          |]
        let resWidth = res |> Array.map Array.length |> Array.max


        let access i j = res.[i] |> (Array.tryItem j >> Option.map XlObj.ofResult >> Option.defaultValue XlObj.Error.xlNA)
        let resRows, resCols, flippedAccess =
          if preferHorizontal && nRows = nCols then
            nRows, resWidth, access
          elif nRows > nCols then
            nRows, resWidth, access
          else
            resWidth, nCols, flip access

        let a = Array2D.init resRows resCols flippedAccess
        return a |> XlObjRange.ofArray2dWithEmpty (XlObj.ofErrorMessage "empty.")
      } |> XlObjRange.ofResult


  type UdfArrayArgWithValue<'T>(name: string, value: xlObj[,], rest: UdfArrayArgBase option, conversion: xlObj -> Result<'T, string>) =
    inherit UdfArrayArgBase(name, value, rest)
    let parsedValueOpt =
      let l1 = value |> Array2D.length1
      let l2 = value |> Array2D.length2
      if l1 = 1 && l2 = 1 then
        value.[0, 0] |> (XlObjParser.withArgName name conversion) |> ArgValue.Scalar
      elif l1 = 1 then
        Array.init l2 (fun j -> value.[0, j] |> (XlObjParser.withArgName name conversion))
        |> ArgValue.Row
      elif l2 = 1 then
        Array.init l1 (fun i -> value.[i, 0] |> (XlObjParser.withArgName name conversion))
        |> ArgValue.Col
      else
        ArgValue.NotCalculated

    member internal uaa.cachedConversion i j =
      match parsedValueOpt with
      | ArgValue.Scalar v -> v
      | ArgValue.Row r -> r.[j]
      | ArgValue.Col c -> c.[i]
      | _ -> value |> getPseudoElt i j |> conversion

    member internal uaa.evalCached f subEval i j =
      result {
        let! fp = subEval f i j
        let! v = uaa.cachedConversion i j
        return fp v
      }


  type ArrayFunctionDefinition<'T1>(name: string, value: xlObj[,], rest: UdfArrayArgBase option, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f (fun f _ _ -> Ok f ) i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalFunction f = uaa.returnArray2d (fun i j -> (uaa.Eval f) i j |> Ok)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalFunction f = uaa.returnArray2d (fun i j -> (uaa.Eval f) i j |> Ok)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6, 'T7>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6, 'T7>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6, 'T7, 'T8>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6, 'T7, 'T8>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6, 'T7, 'T8, 'T9>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6, 'T7, 'T8, 'T9>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgWithValue<'T1>(name, value, rest :> UdfArrayArgBase |> Some, conversion)

    member internal uaa.Eval f i j = uaa.evalCached f rest.Eval i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f) false
    member uaa.EvalArrayFunctionHorizontal f = uaa.returnArray2dFromArrays (uaa.Eval f) true

type ArrayFunctionBuilder() =
  static member Add(name: string, c: xlObj -> _, value: xlObj[,]) =
    ArrayFunction.ArrayFunctionDefinition<_>(name, value, None, c)
