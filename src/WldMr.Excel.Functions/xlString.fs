module WldMr.Excel.String.Basic

open ExcelDna.Integration
open FSharpPlus
open WldMr.Excel.Helpers
open FsToolkit.ErrorHandling


let stringFilter predicate input: obj[,]=
  let f (o: obj) =
    match o with
    | ExcelString s -> s |> predicate
    | _ -> false
  input |> Array2D.map (f >> box)


[<ExcelFunction(Category= "WldMr.String",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string starts with the specified prefix\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringStartsWith
  (
    [<ExcelArgument(Description="The text string value (or range of values) which start is being queried")>]
      input: obj[,],
    [<ExcelArgument(Description="The text string value to be searched for at the start of the input")>]
      prefix: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: obj
  ) =
  monad {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    input
    |> stringFilter (fun s -> s.StartsWith(prefix, ic, System.Globalization.CultureInfo.InvariantCulture))
    |> box
  } |> XlObj.ofResult


[<ExcelFunction(
  Category= "WldMr.String",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string ends with the specified suffix\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringEndsWith
  (
    [<ExcelArgument(Description="The text string value (or range of values) which end is being queried")>]
      input: obj[,],
    [<ExcelArgument(Description="The text string value to be searched for at the end of the input")>]
      suffix: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: obj
  ) =
  monad {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    input
    |> stringFilter (fun s -> s.EndsWith(suffix, ic, System.Globalization.CultureInfo.InvariantCulture))
    |> box
  } |> XlObj.ofResult


[<ExcelFunction(Category= "WldMr.String",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string contains the specified substring\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringContains
  (
    [<ExcelArgument(Description="The text string value (or range of values) which is being queried")>]
      input: obj[,],
    [<ExcelArgument(Description="The text string value to be searched within the input")>]
      subString: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: obj
  ) =
  monad {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    input
    |> stringFilter (fun s -> if ic then s.ToLowerInvariant().Contains(subString.ToLowerInvariant()) else s.Contains(subString))
    |> box
  } |> XlObj.ofResult
