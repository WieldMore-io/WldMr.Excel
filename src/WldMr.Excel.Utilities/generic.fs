namespace WldMr.Excel.Utilities


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

