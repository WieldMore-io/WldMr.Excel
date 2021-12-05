namespace WldMr.Excel


[<AutoOpen>]
module XlObjArray =

  [<RequireQualifiedAccess>]
  module XlObj =

    [<RequireQualifiedAccess>]
    type TrimMode = MissingEmpty | MissingEmptyStringEmpty

    [<RequireQualifiedAccess>]
    module Array =


      /// <summary>
      /// Drops trailing elements that do meet the trimMode predicate
      /// This might return an empty array
      /// </summary>
      let trim trimMode (x: objCell[]) =
        let trim_ pred x =
          match x |> Array.tryFindIndexBack pred with
          | None -> [||]
          | Some n -> x.[..n]
        let predicate =
          match trimMode with
          | TrimMode.MissingEmpty -> (function | ExcelMissing _ | ExcelEmpty _ | _ -> true)
          | TrimMode.MissingEmptyStringEmpty -> (function | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> false | _ -> true)
        trim_ predicate x

      /// <summary>
      /// </summary>
      let toFloatArray (o: objCell[]) =
        let floats = Array.zeroCreate o.Length
        let mutable error = None
        for i = 0 to o.Length - 1 do
          match o.[i] with
          | ExcelNum f -> floats.[i] <- f
          | ExcelString s ->
              match System.Double.TryParse s with
              | false, _ -> error <- Some $"Could not parse {s} as a number."
              | true, f -> floats.[i] <- f
          | _ -> error <- Some $"Could not parse {o.[i]} as a number."
        match error with
        | Some e -> Error e
        | None -> Ok floats

      /// <summary>
      /// </summary>
      let toStringArray (o: objCell[]) =
        let strings = Array.zeroCreate o.Length
        let mutable error = None
        for i = 0 to o.Length - 1 do
          match o.[i] with
          | ExcelString s ->
              strings.[i] <- s
          | _ -> error <- Some $"Could not parse {o.[i]} as a string."
        match error with
        | Some e -> Error e
        | None -> Ok strings

      /// <summary>
      /// </summary>
      let toIntArray (o: objCell[]) =
        o
        |> toFloatArray
        |> Result.map (Array.map int)


    /// <summary>
    /// Returns a column array from a sequence
    /// </summary>
    [<RequireQualifiedAccess>]
    module Column =
      let ofSeqWithEmpty (emptyVal: objCell) (r: seq<objCell>) =
        let v = r |> Array.ofSeq
        if v.Length = 0 then
          emptyVal |> Array2D.create 1 1
        else
          Array2D.init v.Length 1 (fun i _ -> v.[i])

    /// <summary>
    /// Returns a row array from a sequence
    /// </summary>
    [<RequireQualifiedAccess>]
    module Row =
      let ofSeqWithEmpty (emptyVal: objCell) (r: seq<objCell>) =
        let v = r |> Array.ofSeq
        if v.Length = 0 then
          emptyVal |> Array2D.create 1 1
        else
          Array2D.init 1 v.Length (fun _ j -> v.[j])
