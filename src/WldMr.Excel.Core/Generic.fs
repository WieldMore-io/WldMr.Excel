namespace WldMr.Excel.Core

open WldMr.Excel

module Generic =
  open TypeShape.Core

  let tupleToArray<'T> () : 'T -> xlObj[] =
    match shapeof<'T> with
    | Shape.Tuple (:? ShapeTuple<'T> as shape) ->
        let mkElemToObj (shape : IShapeMember<'T>) =
           shape.Accept { new IMemberVisitor<'T, 'T -> xlObj> with
               member _.Visit (shape : ShapeMember<'T, 'Field>) =
                  shape.Get >> unbox<'Field -> xlObj>(fun (o:'Field) -> o |> XlObj.Unsafe.ofObj)
               }

        let elemToObjs : ('T -> xlObj) [] = shape.Elements |> Array.map mkElemToObj

        fun (r:'T) -> elemToObjs |> Array.map (fun ep -> ep r)

    | _ -> failwithf "unsupported type '%O'" typeof<'T>

  let tupleWidth<'T> () : 'T [] -> int =
    match shapeof<'T> with
    | Shape.Tuple (:? ShapeTuple<'T> as shape) ->
        fun (r:'T []) -> shape.Elements.Length
    | _ -> failwithf "unsupported type '%O'" typeof<'T>

