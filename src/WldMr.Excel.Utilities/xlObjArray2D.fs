namespace WldMr.Excel.Utilities


[<AutoOpen>]
module XlObjArray2D =

  [<RequireQualifiedAccess>]
  module XlObj =

    [<RequireQualifiedAccess>]
    module Array2D =

      let private trimPrivate pred (x:obj[,]) =
        let x0, x1 = x |> XlObj.getSize

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
      let trim trimMode (x: obj[,]) =
        let predicate =
          match trimMode with
          | XlObj.TrimMode.MissingEmpty -> (function | ExcelMissing _ | ExcelEmpty _ | _ -> true)
          | XlObj.TrimMode.MissingEmptyStringEmpty -> (function | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> false | _ -> true)
        trimPrivate predicate x

