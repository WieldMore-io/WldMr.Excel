module WldMr.Excel.Functions.Format

open ExcelDna.Integration
open FsToolkit.ErrorHandling

open WldMr.Excel
open WldMr.Excel.Functions



// the usability of this function is not great
// - error reporting is not up to par
// - it would natural to pass an array of values and an array of types
// - or maybe to arrayfy the function on inputs
[<ExcelFunction(Category = "WldMr Text",
                Description = """Formats an interpolated string as in the .Net world"
eg xlFormat("{0:dd-mmm-yy}", A1, "d")
eg xlFormat("{0:0.00}", A1, "f")
See help link for more details about the syntax""",
                HelpTopic = "https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types")>]
let xlFormatA
  (
    [<ExcelArgument(Description="Format string")>]
    s: string,
    o1: objCell, t1: objCell,
    o2: objCell, t2: objCell,
    o3: objCell, t3: objCell,
    o4: objCell, t4: objCell,
    o5: objCell, t5: objCell,
    o6: objCell, t6: objCell,
    o7: objCell, t7: objCell,
    o8: objCell, t8: objCell
  ): objCell =
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
        | Ok x -> $"Invalid format specifier '{x}'. Use d, i, f, or s." |> Error
        | Error e -> $"Invalid format specifier. Use d, i, f, or s. {e}" |> Error

  let formatString s args =
    try
      System.String.Format(s, List.toArray args) |> XlObj.ofString |> Ok
    with
      | :? System.FormatException as e -> "Incorrect format string." |> Error
      | e -> $"{e.Message} ({e.GetType()})" |> Error

  result {
    let! args =
      ([ o1; o2; o3; o4; o5; o6; o7; o8 ],
       [ t1; t2; t3; t4; t5; t6; t7; t8 ])
      ||> List.map2 convertXlObj
      |> List.sequenceResultM

    return! formatString s args
  } |> XlObj.ofResult
