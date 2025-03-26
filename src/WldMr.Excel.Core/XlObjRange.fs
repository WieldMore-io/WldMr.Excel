namespace WldMr.Excel

open System


[<RequireQualifiedAccess>]
module XlObjRange =

  let getSize (a: xlObj[,]): int * int =
    match a.GetLength 0, a.GetLength 1 with
    | 0, _ | _, 0 -> 0, 0
    | 1, 1 when a.[0, 0] = XlObj.xlMissing -> 0, 0
    | ls -> ls

  let isMissing (a: xlObj[,]): bool =
    (getSize a) = (0, 0)

[<AutoOpen>]
module OfConversions =
  [<RequireQualifiedAccess>]
  module XlObjRange =

    let ofCell (o: xlObj) = Array2D.create 1 1 o

    let ofResult (t: Result<xlObj[,], string>): xlObj[,] =
      match t with
      | Ok v -> v
      | Error err -> err |> XlObj.ofErrorMessage |> ofCell

    let ofValidation (t: Result<xlObj[,], string list>): xlObj[,] =
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
    let ofArray2d(a: xlObj[,]): xlObj[,] =
      if a.Length = 0 then
        XlObj.ofErrorMessage "empty range." |> ofCell
      else
        a

    /// <summary>
    /// </summary>
    let ofArray2dWithEmpty (emptyValue: xlObj) (a: xlObj[,]): xlObj[,] =
      if a.Length = 0 then
        emptyValue |> ofCell
      else
        a

    let ofRows(rows: #seq<#seq<xlObj>>) =
      let maxLength = rows |> Seq.map Seq.length |> Seq.max
      let nRows = rows |> Seq.length
      let a = Array2D.create nRows maxLength (XlObj.ofString "")

      rows
      |> Seq.iteri (fun i ->
        Seq.iteri (fun j item ->
          a.[i, j] <- item
        )
      )
      a

    let ofRowsWithEmpty (emptyVal: xlObj) (rows: #seq<#seq<xlObj>>) =
      let nRows = rows |> Seq.length
      if nRows = 0 then emptyVal |> ofCell else

      let maxLength = rows |> Seq.map Seq.length |> Seq.max
      let a = Array2D.create nRows maxLength (XlObj.ofString "")

      rows
      |> Seq.iteri (fun i ->
        Seq.iteri (fun j item ->
          a.[i, j] <- item
        )
      )
      a


[<AutoOpen>]
module ToConversions =
  [<RequireQualifiedAccess>]
  module XlObjRange =

    /// <summary>
    /// </summary>
    let toFloatArray (o: xlObj[]) =
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
    let toStringArray (o: xlObj[]) =
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
    /// Empty cells are treated as empty strings
    /// Number cells are converted to string
    /// Error cells still produce an Error
    /// </summary>
    let toStringArrayPermissive (o: xlObj[]) =
      let strings = Array.zeroCreate o.Length
      let mutable error = None
      for i = 0 to o.Length - 1 do
        match o.[i] with
        | ExcelEmpty _ -> strings.[i] <- ""
        | ExcelNum f when abs ( f - (f |> int |> float) ) < 1e-8 -> strings.[i] <- f |> int |> string
        | ExcelNum f -> strings.[i] <- f |> string
        | ExcelString s -> strings.[i] <- s
        | _ -> error <- Some $"Could not parse {o.[i]} as a string."
      match error with
      | Some e -> Error e
      | None -> Ok strings

    /// <summary>
    /// </summary>
    let toIntArray (o: xlObj[]) =
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
      let ofSeqWithEmpty (emptyVal: xlObj) (r: #seq<xlObj>) =
        let v = r |> Array.ofSeq
        if v.Length = 0 then
          emptyVal |> XlObjRange.ofCell
        else
          Array2D.init v.Length 1 (fun i _ -> v.[i])

      let ofSeq r =
        ofSeqWithEmpty XlObj.Error.xlNA r

    /// <summary>
    /// Returns a row array from a sequence
    /// </summary>
    [<RequireQualifiedAccess>]
    module Row =
      let ofSeqWithEmpty (emptyVal: xlObj) (r: #seq<xlObj>) =
        let v = r |> Array.ofSeq
        if v.Length = 0 then
          emptyVal |> XlObjRange.ofCell
        else
          Array2D.init 1 v.Length (fun _ j -> v.[j])

      let ofSeq r =
        ofSeqWithEmpty XlObj.Error.xlNA r

    /// <summary>
    /// Returns a range from a sequence of rows
    /// </summary>
    [<RequireQualifiedAccess>]
    module Rows =
      let ofSeqWithEmpty (emptyVal: xlObj) (r: #seq<#seq<xlObj>>) =
        if r |> Seq.length = 0 then
          emptyVal |> XlObjRange.ofCell
        else
          array2D r

      let ofSeq r =
        ofSeqWithEmpty XlObj.Error.xlNA r


    /// <summary>
    /// Returns a range from a sequence of rows
    /// </summary>
    [<RequireQualifiedAccess>]
    module Columns =
      let ofSeqWithEmpty (emptyVal: xlObj) (r: #seq<#seq<xlObj>>) =
        let rowsR = Rows.ofSeqWithEmpty emptyVal r
        Array2D.init (rowsR.GetLength 1) (rowsR.GetLength 0) (fun i j -> rowsR.[j, i])

      let ofSeq r =
        ofSeqWithEmpty XlObj.Error.xlNA r
