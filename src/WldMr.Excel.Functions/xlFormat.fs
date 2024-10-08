module WldMr.Excel.Functions.Format

open ExcelDna.Integration
open FsToolkit.ErrorHandling
open WldMr.Excel


// the usability of this function is not great
// - error reporting is not up to par
// - it would natural to pass an array of values and an array of types
// - or maybe to arrayify the function on inputs
[<ExcelFunction(Category = "WldMr Text",
                Description = """Formats an interpolated string as in the .Net world"
eg xlFormat("{0:dd-mmm-yy}", A1, "d")
eg xlFormat("{0:0.00}", A1, "f")
See help link for more details about the syntax""",
                HelpTopic = "https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types")>]
let xlFormatA
  (
    [<ExcelArgument(Description="Format string")>]
    s: xlObj[,],
    o1: xlObj[,], t1: xlObj[,],
    o2: xlObj[,], t2: xlObj[,],
    o3: xlObj[,], t3: xlObj[,],
    o4: xlObj[,], t4: xlObj[,],
    o5: xlObj[,], t5: xlObj[,],
    o6: xlObj[,], t6: xlObj[,]
  ): xlObj[,] =
  let convertXlObj (o: xlObj) (t: xlObj) =
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
      | :? System.FormatException as e -> $"Incorrect format string: {e.Message}" |> Error
      | e -> $"{e.Message} ({e.GetType()})" |> Error

  let scalarF
    (s: string)
    (o1: xlObj) (t1: xlObj)
    (o2: xlObj) (t2: xlObj)
    (o3: xlObj) (t3: xlObj)
    (o4: xlObj) (t4: xlObj)
    (o5: xlObj) (t5: xlObj)
    (o6: xlObj) (t6: xlObj)
    =
    result {
      let! args =
        ([ o1; o2; o3; o4; o5; o6],
         [ t1; t2; t3; t4; t5; t6])
        ||> List.map2 convertXlObj
        |> List.sequenceResultM

      return! formatString s args
    }

  ArrayFunctionBuilder
    .Add("formatString", XlObj.toString, s)
    .Add("o1", Ok, o1).Add("t1", Ok, t1)
    .Add("o2", Ok, o2).Add("t2", Ok, t2)
    .Add("o3", Ok, o3).Add("t3", Ok, t3)
    .Add("o4", Ok, o4).Add("t4", Ok, t4)
    .Add("o5", Ok, o5).Add("t5", Ok, t5)
    .Add("o6", Ok, o6).Add("t6", Ok, t6)
    .EvalFunction scalarF
  |> FunctionCall.eval
