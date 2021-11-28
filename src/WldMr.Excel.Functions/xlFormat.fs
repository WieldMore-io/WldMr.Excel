module WldMr.Excel.String.Format

open ExcelDna.Integration
open FsToolkit.ErrorHandling

open WldMr.Excel.Utilities
open WldMr

[<ExcelFunction(Category = "WldMr Text",
                Description = """Formats an interpolated string as in the .Net world"
\teg xlFormat("dd-mmm-yy", A1, "d")
See help link for more details about the syntax""",
                HelpTopic = "https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types")>]
let xlFormatA
  (
    s: string,
    o1: objCell, t1: objCell,
    o2: objCell, t2: objCell,
    o3: objCell, t3: objCell,
    o4: objCell, t4: objCell,
    o5: objCell, t5: objCell,
    o6: objCell, t6: objCell,
    o7: objCell, t7: objCell,
    o8: objCell, t8: objCell
  ) =
  let convertXlObj (o: objCell) (t: objCell) =
    match o, t with
    | ExcelMissing _, ExcelMissing _ -> "" :> obj |> Ok
    | _, _ ->
        match t |> XlObj.toString |> Result.map String.toLower with
        | Ok "d" ->
            o |> XlObj.toFloat
            |> Result.map (System.DateTime.FromOADate >> box)
        | Ok "i" -> o |> XlObj.toFloat |> Result.map (int >> box)
        | Ok "f" -> o |> XlObj.toFloat |> Result.map box
        | Ok "s" -> o |> XlObj.toString |> Result.map box
        | _ ->
            "Invalid format specifier. Use d, i, f, or s."
            |> Error
        |> Validation.ofResult

  validation {
    let! args =
      ([ o1; o2; o3; o4; o5; o6; o7; o8 ],
       [ t1; t2; t3; t4; t5; t6; t7; t8 ])
      ||> List.map2 convertXlObj
      |> List.sequenceValidationA
      |> Result.map List.toArray

    return! (s, args)
      |> Result.protect (System.String.Format >> box >> (~%))
      |> Result.mapError (fun err -> [ $"{err}" ])
  } |> XlObj.ofValidation
