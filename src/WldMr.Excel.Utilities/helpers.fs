namespace WldMr.Excel.Helpers

open ExcelDna.Integration
open System
open FSharpPlus


[<AutoOpen>]
module SingletonValue =
  let objEmpty = ExcelEmpty.Value :> obj
  let objMissing = ExcelMissing.Value :> obj
  let objNA = ExcelError.ExcelErrorNA :> obj
  let objName = ExcelError.ExcelErrorName :> obj
  let objGettingData = ExcelError.ExcelErrorGettingData :> obj


module Result =
  /// The error value is prefixed by "#!Error! "
  let inline toExcel r = r |> Result.defaultWith (fun e -> $"#Error! {e}" :> obj)
  let mapArgError errMsg = Result.mapError (fun e -> [$"Arg '{errMsg}': {e}"])


module Array2D =
  /// flattens the array as a list of rows, each row being a list
  /// simplify further processing (although potentially slower)
  let flattenArray array2d =
    [ for x in [ 0 .. (Array2D.length1 array2d) - 1 ] do
      [ for y in [ 0 .. (Array2D.length2 array2d) - 1 ] do
        yield array2d.[x, y] ] ]


[<AutoOpen>]
module ActivePattern =
  let (|ExcelMissingRange|_|) (input: obj[,]) =
    if input.GetLength 0 = 1 && input.GetLength 1 = 1 then
      match input.[0, 0] with
      | :? ExcelMissing as m -> Some m
      | _ -> None
    else
      None

  let (|ExcelMissing|_|) (input: obj) =
    match input with
    | :? ExcelMissing as m -> Some m
    | _ -> None

  let (|ExcelEmpty|_|) (input: obj) =
    match input with
    | :? ExcelEmpty as m -> Some m
    | _ -> None

  let (|ExcelString|_|) (input: obj) =
    match input with
    | :? string as s -> Some s
    | _ -> None

  let (|ExcelError|_|) (input: obj) =
    match input with
    | :? ExcelError as e -> Some e
    | _ -> None

  let (|ExcelBool|_|) (input: obj) =
    match input with
    | :? bool as s -> Some s
    | _ -> None

  let (|ExcelNum|_|) (input: obj) =
    match input with
    | :? float as s -> Some s
    | _ -> None


module XlObj =
  /// Defaults the value if the input is Missing, Empty or ""
  let defaultWith defaultFun (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultFun () |> box
    | o -> o

  let defaultValue v (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> v |> box
    | o -> o

  /// Tries to extract an int out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toInt (o: obj) =
    match o with | :? float as f -> f |> int |> Ok | _ -> "Expected a number" |> Error

  /// Tries to extract a float out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toFloat (o: obj) =
    match o with | :? float as f -> f |> Ok | _ -> "Expected a number" |> Error

  /// Tries to extract a datetime out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toDate (o: obj) =
    match o with | :? float as f -> f |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  /// Tries to extract a date without time out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toDateNoTime (o: obj) =
    match o with | :? float as f -> f |> int |> float |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  /// try to extract a string without time out of excel cell value
  /// does not attempt any conversion (the original excel value must be a string)
  let toString (o: obj) =
    match o with | :? string as s -> s |> Ok | _ -> "Expected a string" |> Error

  /// force converts an Excel obj input to a boolean option
  /// - missing, empty, or empty-string values are treated as None (!)
  /// - Errors, exact 0 are treated as False (as expected)
  /// - Strings are treated as False (debatable)
  /// - non-zero numbers are treated as True
  let toBoolOption (o:obj) =
    match o with
    | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> None |> Ok
    | ExcelError _ | ExcelNum 0.0 -> Some false |> Ok
    | ExcelNum _ -> Some true |> Ok
    | ExcelBool b -> Some b |> Ok
    | _ ->  Some false |> Ok


  let toBool (o:obj) =
    match o with
    | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> false |> Ok
    | ExcelError _ | ExcelNum 0.0 -> false |> Ok
    | ExcelNum _ -> true |> Ok
    | ExcelBool b -> b |> Ok
    | _ ->  "Expected a boolean" |> Error


  let toBoolWithDefault d (o:obj) =
    match o with
    | ExcelEmpty _ -> false |> Ok
    | ExcelMissing _ -> d |> Ok
    | ExcelNum 0.0 -> false |> Ok
    | ExcelNum _ -> true |> Ok
    | ExcelBool b -> b |> Ok
    | ExcelString s when s.ToLower() = "true" -> true |> Ok
    | ExcelString s when s.ToLower() = "false" -> false |> Ok
    | _ ->  "Expected a boolean" |> Error


  /// boxes the float
  /// if its value is NaN, it is replaced by #N/A!
  /// does not attempt any conversion (the original excel value must be a string)
  let ofFloat f =
    if Double.IsNaN f then
      ExcelError.ExcelErrorNA :> obj
    else
      f |> box

  /// Summarize the number of errors and returns them prefixed by "#Error!""
  let ofValidation (t: Result<obj, string list>): obj =
    let errorMessage errors =
      let sep = "; "
      match errors with
      | [] -> "Unexpected error"
      | x::[] -> $"#Error! {x}"
      | xs -> $"#Error! {xs.Length} errors: {String.Join(sep, xs)}"

    t |> Result.either id (errorMessage >> box)

  let ofResult<'E> (t: Result<obj, 'E>): obj =
    t |> Result.either id (fun err -> $"#Error! {err}" :> obj)


module ExcelAsync =
  let wrap<'T> (funName: string) (hash: int) (errMsg: obj) (waitMsg: obj) (f: unit -> obj): obj =
    let asyncResult = ExcelAsyncUtil.Run(funName, hash, new ExcelFunc(f))
    match asyncResult with
    | null -> errMsg
    | ExcelError(ExcelError.ExcelErrorNA) -> waitMsg
    | :? System.Tuple<string, string> as tu ->
        let t1, t2 = tu
        (if t1 = null then t2 else t1) |> box
    | _ -> asyncResult

  let excelObserve functionName parameters observable =
    let obsSource =
        ExcelObservableSource(
            fun () ->
            { new IExcelObservable with
                member __.Subscribe observer =
                    // Subscribe to the F# observable
                    Observable.subscribe (fun value -> observer.OnNext (value)) observable
            })
    ExcelAsyncUtil.Observe (functionName, parameters, obsSource)


module XlArray =
  let columnFromSeq r =
    let v = r |> Array.ofSeq
    Array2D.init v.Length 1 (fun i j -> v.[i])


module Generic =
  open TypeShape.Core

  let tupleToArray<'T> () : 'T -> obj[] =
    match shapeof<'T> with
    | Shape.Tuple (:? ShapeTuple<'T> as shape) ->
        let mkElemToObj (shape : IShapeMember<'T>) =
           shape.Accept { new IMemberVisitor<'T, 'T -> obj> with
               member _.Visit (shape : ShapeMember<'T, 'Field>) =
                  shape.Get >> unbox<'Field -> obj>(fun (o:'Field) -> o |> box)
               }

        let elemToObjs : ('T -> obj) [] = shape.Elements |> Array.map mkElemToObj

        fun (r:'T) -> elemToObjs |> Array.map (fun ep -> ep r)

    | _ -> failwithf "unsupported type '%O'" typeof<'T>

  let tupleWidth<'T> () : 'T [] -> int =
    match shapeof<'T> with
    | Shape.Tuple (:? ShapeTuple<'T> as shape) ->
        fun (r:'T []) -> shape.Elements.Length
    | _ -> failwithf "unsupported type '%O'" typeof<'T>


