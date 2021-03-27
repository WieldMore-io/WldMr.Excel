module WldMr.Excel.String.Format

open ExcelDna.Integration
open FSharpPlus
open FsToolkit.ErrorHandling

open WldMr.Excel.Helpers

[<ExcelFunction(Category= "WldMr.String",
  Description= """Format an interpolated string
eg xlFormat("dd-mmm-yy", A1, "d")
https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types""",
  HelpTopic="https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types"
)
  >]
let xlFormatA
  (
    s:string,
    o1:obj, t1:obj,
    o2: obj, t2:obj,
    o3: obj, t3:obj,
    o4: obj, t4:obj,
    o5: obj, t5:obj,
    o6: obj, t6:obj,
    o7: obj, t7:obj,
    o8: obj, t8:obj
  ) =
  let formatValue (o: obj) (t: obj) =
    match (o, t) with
    | (ExcelMissing _, ExcelMissing _) -> "" :> obj |> Result.Ok
    | (_, _) ->
      match t |> XlObj.toString |>> String.toLower with
      | Ok "d" -> o |> XlObj.toFloat |> map (System.DateTime.FromOADate >> box)
      | Ok "i" -> o |> XlObj.toFloat |>> (int >> box)
      | Ok "f" -> o |> XlObj.toFloat |>> box
      | Ok "s" -> o |> XlObj.toString |>> box
      | _ -> "Invalid format specifier. Use d, i, f, or s." |> Error
      |> Validation.ofResult

  let resToFlatten = validation {
    let! args =
      ([o1; o2; o3; o4; o5; o6; o7; o8],
       [t1; t2; t3; t4; t5; t6; t7; t8])
      ||> List.map2 formatValue
      |> List.sequenceValidationA
      |>> List.toArray

    return! Result.protect (fun () ->
      System.String.Format(s, args) :> obj) ()
      |> Result.mapError (fun err -> [$"{err}"]
    )
  }
  let res = resToFlatten // |> Result.flatten
  res |> XlObj.ofResult
