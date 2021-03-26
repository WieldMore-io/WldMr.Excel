namespace WldMr.Excel.Helpers

open ExcelDna.Integration
open System
open FSharpPlus


module Result =
  let inline toExcel r = r |> Result.defaultWith (fun e -> $"#Error! {e}" :> obj)


module Array2D = 
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
  let toInt (o: obj) = 
    match o with | :? float as f -> f |> int |> Ok | _ -> "Expected a number" |> Error

  let toFloat (o: obj) = 
    match o with | :? float as f -> f |> Ok | _ -> "Expected a number" |> Error

  let toDate (o: obj) = 
    match o with | :? float as f -> f |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  let toDateNoTime (o: obj) = 
    match o with | :? float as f -> f |> int |> float |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  let toString (o: obj) = 
    match o with | :? string as s -> s |> Ok | _ -> "Expected a string" |> Error

  let toBoolOption (o:obj) =
    match o with
    | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> None |> Ok
    | ExcelError _ | ExcelNum 0.0 -> Some false |> Ok
    | ExcelNum _ -> Some true |> Ok
    | ExcelBool b -> Some b |> Ok
    | _ ->  Some false |> Ok


  let ofFloat f =
    if Double.IsNaN f then 
      ExcelDna.Integration.ExcelError.ExcelErrorNA :> obj
    else
      f |> box

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
  let wrap<'T> (funName: string) (hashInt: int) (errorMessage: string) (waitingMessage: string) (f: unit -> obj): obj =
    let asyncResult = ExcelAsyncUtil.Run(funName, hashInt, new ExcelFunc(f))
    match asyncResult with
    | null -> errorMessage |> box
    | ExcelError(ExcelError.ExcelErrorNA) -> waitingMessage |> box
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


