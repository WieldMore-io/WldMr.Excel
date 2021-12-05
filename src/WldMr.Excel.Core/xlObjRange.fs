namespace WldMr.Excel

open System


[<RequireQualifiedAccess>]
module XlObjRange =

  let getSize (a: objCell[,]): int * int =
    match a.GetLength 0, a.GetLength 1 with
    | 0, _ | _, 0 -> 0, 0
    | 1, 1 when a.[0, 0] = XlObj.xlMissing -> 0, 0
    | ls -> ls


[<AutoOpen>]
module OfConversions =
  [<RequireQualifiedAccess>]
  module XlObjRange =

    let ofCell (o: objCell) = Array2D.create 1 1 o

    let ofResult (t: Result<objCell[,], string>): objCell[,] =
      match t with
      | Ok v -> v
      | Error err -> err |> XlObj.ofErrorMessage |> ofCell

    let ofValidation (t: Result<objCell[,], string list>): objCell[,] =
      let errorMessage errors =
        let sep = "; "
        match errors with
        | [] -> "Unexpected error"
        | [ x ] -> x
        | xs -> $"{xs.Length} errors: {String.Join(sep, xs)}"

      match t with
      | Ok v -> v
      | Error e -> e |> errorMessage |> XlObj.ofErrorMessage |> ofCell

    /// <summary>
    /// </summary>
    let ofArray2d(a: objCell[,]): objCell[,] =
      if a.Length = 0 then
        XlObj.ofErrorMessage "empty range." |> ofCell
      else
        a

    /// <summary>
    /// </summary>
    let ofArray2dWithEmpty (emptyValue: objCell) (a: objCell[,]): objCell[,] =
      if a.Length = 0 then
        emptyValue |> ofCell
      else
        a


[<AutoOpen>]
module ToConversions =
  [<RequireQualifiedAccess>]
  module XlObjRange =

    /// <summary>
    /// </summary>
    let toFloatArray (o: objCell[]) =
      let floats = Array.zeroCreate o.Length
      let mutable error = None
      for i = 0 to o.Length - 1 do
        match o.[i] with
        | ExcelNum f -> floats.[i] <- f
        | ExcelString s ->
            match Double.TryParse s with
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

[<AutoOpen>]
module RowColumn =
  [<RequireQualifiedAccess>]
  module XlObjRange =
    /// <summary>
    /// Returns a column array from a sequence
    /// </summary>
    [<RequireQualifiedAccess>]
    module Column =
      let ofSeqWithEmpty (emptyVal: objCell) (r: seq<objCell>) =
        let v = r |> Array.ofSeq
        if v.Length = 0 then
          emptyVal |> XlObjRange.ofCell
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
          emptyVal |> XlObjRange.ofCell
        else
          Array2D.init 1 v.Length (fun _ j -> v.[j])

