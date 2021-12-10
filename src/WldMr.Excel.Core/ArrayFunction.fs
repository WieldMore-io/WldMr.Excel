namespace WldMr.Excel

open FsToolkit.ErrorHandling
open WldMr.Excel


[<RequireQualifiedAccess>]
module ArrayFunction =

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

    member internal uaa.returnArray2dFromArrays f () =
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
          if nRows > nCols then
            nRows, resWidth, access
          else
            resWidth, nCols, flip access

        let a = Array2D.init resRows resCols flippedAccess
        return a |> XlObjRange.ofArray2dWithEmpty (XlObj.ofErrorMessage "empty.")
      } |> XlObjRange.ofResult


  type ArrayFunctionDefinition<'T1>(name: string, value: xlObj[,], rest: UdfArrayArgBase option, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest)
    member private uaa.Conversion = conversion

    member internal uaa.Eval f i j = uaa.eval f (fun f _ _ -> Ok f ) uaa.Conversion i j

    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalFunction f = uaa.returnArray2d (fun i j -> (uaa.Eval f) i j |> Ok)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalFunction f = uaa.returnArray2d (fun i j -> (uaa.Eval f) i j |> Ok)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

    member uaa.Add(name: string, c: xlObj -> _, value: xlObj[,]) =
      ArrayFunctionDefinition<_, _, _, _, _, _, _>(name, value, uaa, c)

  and ArrayFunctionDefinition<'T1, 'T2, 'T3, 'T4, 'T5, 'T6, 'T7>(name: string, value: xlObj[,], rest: ArrayFunctionDefinition<'T2, 'T3, 'T4, 'T5, 'T6, 'T7>, conversion: xlObj -> Result<'T1, string> ) =
    inherit UdfArrayArgBase(name, value, rest :> UdfArrayArgBase |> Some)

    member private uaa.Conversion = conversion
    member internal uaa.Eval f i j = uaa.eval f rest.Eval uaa.Conversion i j
    member uaa.EvalFunction f = uaa.returnArray2d (uaa.Eval f)
    member uaa.EvalArrayFunction f = uaa.returnArray2dFromArrays (uaa.Eval f)

type ArrayFunctionBuilder() =
  static member Add(name: string, c: xlObj -> _, value: xlObj[,]) =
    ArrayFunction.ArrayFunctionDefinition<_>(name, value, None, c)
