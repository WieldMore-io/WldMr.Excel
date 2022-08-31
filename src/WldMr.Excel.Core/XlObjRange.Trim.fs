namespace WldMr.Excel


[<AutoOpen>]
module XlObjRangeTrimOps =

  [<RequireQualifiedAccess>]
  module XlObjRange =

    [<RequireQualifiedAccess>]
    type TrimMode = MissingEmpty | MissingEmptyStringEmpty

    module TrimMode =
      let predicate trimMode =
        match trimMode with
        | TrimMode.MissingEmpty -> (function | ExcelMissing _ | ExcelEmpty _ -> false | _ -> true)
        | TrimMode.MissingEmptyStringEmpty -> (function | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> false | _ -> true)


    /// <summary>
    /// Drops trailing elements that do meet the trimMode predicate
    /// This might return an empty array
    /// </summary>
    let trimArray trimMode (x: xlObj[]) =
      let trim_ pred x =
        match x |> Array.tryFindIndexBack pred with
        | None -> [||]
        | Some n -> x.[..n]
      trim_ (trimMode |> TrimMode.predicate) x

    [<RequireQualifiedAccess>]
    module Array2DInternal =

      let trimPrivate pred (x:xlObj[,]) =
        let x0, x1 = x |> XlObjRange.getSize

        let lastRow =
          seq { for i in x0-1 .. -1 .. 0 do yield x.[i, *] }
          |> Seq.tryFindIndex (Array.exists pred)
          |> Option.defaultValue x0
          |> (-) (x0 - 1)

        let lastCol =
          seq { for j in x1-1 .. -1 .. 0 do yield x.[0..lastRow, j] }
          |> Seq.tryFindIndex (Array.exists pred)
          |> Option.defaultValue x1
          |> (-) (x1 - 1)

        x.[0..lastRow, 0..lastCol]

    /// <summary>
    /// Drops trailing rows and columns that do meet the trimMode predicate
    /// This might return an empty array
    /// </summary>
    let trimRange trimMode (x: xlObj[,]) =
      Array2DInternal.trimPrivate (trimMode |> TrimMode.predicate) x

