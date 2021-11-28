module WldMr.Excel.String.Basic

open WldMr.Excel.Utilities
open WldMr
open FsToolkit.ErrorHandling
open ExcelDna.Integration


let stringFilter (predicate: string -> bool) (input: objCell[,]): objCell[,] =
  let f (o: objCell) =
    match o with
    | ExcelString s -> s |> predicate
    | _ -> false
  input |> Array2D.map (f >> box >> (~%))


open System.Text.RegularExpressions

let regexFilter regex ignoreCase input: Result<objCell[,], string> =
  let regexOptions =
    if ignoreCase then
      RegexOptions.Compiled ||| RegexOptions.CultureInvariant ||| RegexOptions.IgnoreCase
    else
      RegexOptions.Compiled ||| RegexOptions.CultureInvariant
  result {
    let! r =
      Result.protect (fun () -> new System.Text.RegularExpressions.Regex(regex, regexOptions)) ()
      |> Result.mapError (fun ex -> ex.ToString())

    let f (o: objCell) =
      match o with
      | ExcelString s -> s |> r.IsMatch
      | _ -> false
    return input |> Array2D.map (f >> box >> (~%))
  }

[<ExcelFunction(Category= "WldMr Text",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string starts with the specified prefix\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringStartsWith
  (
    [<ExcelArgument(Description="The text string value (or range of values) which start is being queried")>]
      input: objCell[,],
    [<ExcelArgument(Description="The text string value to be searched for at the start of the input")>]
      prefix: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: objCell,
    [<ExcelArgument(Description="If TRUE, 'prefix' is a regular expression, if FALSE or omitted, 'prefix' is a literal")>]
      useRegex: objCell
  ) =
  result {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    let! ur = useRegex |> XlObj.toBoolWithDefault false
    if ur then
      let adjPrefix = if prefix.StartsWith "^" then prefix else "^" + prefix
      return! input |> regexFilter adjPrefix ic
    else
      return
        input
        |> stringFilter (fun s -> s.StartsWith(prefix, ic, System.Globalization.CultureInfo.InvariantCulture))
  } |> XlObjRange.ofResult


[<ExcelFunction(
  Category= "WldMr Text",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string ends with the specified suffix\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringEndsWith
  (
    [<ExcelArgument(Description="The text string value (or range of values) which end is being queried")>]
      input: objCell[,],
    [<ExcelArgument(Description="The text string value to be searched for at the end of the input")>]
      suffix: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: objCell,
    [<ExcelArgument(Description="If TRUE, 'suffix' is a regular expression, if FALSE or omitted, 'suffix' is a literal")>]
      useRegex: objCell
  ) =
  result {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    let! ur = useRegex |> XlObj.toBoolWithDefault false
    if ur then
      let adjPrefix = if suffix.EndsWith "$" then suffix else suffix + "$"
      return! input |> regexFilter adjPrefix ic
    else
      return
        input
        |> stringFilter (fun s -> s.EndsWith(suffix, ic, System.Globalization.CultureInfo.InvariantCulture))
  } |> XlObjRange.ofResult


[<ExcelFunction(Category= "WldMr Text",
  IsThreadSafe=true,
  Description=
    "Returns TRUE if the text string contains the specified substring\r\n" +
    "This function also operates on arrays\r\n" +
    "Returns FALSE for any non text input\r\n"
)>]
let xlStringContains
  (
    [<ExcelArgument(Description="The text string value (or range of values) which is being queried")>]
      input: objCell[,],
    [<ExcelArgument(Description="The text string value to be searched within the input")>]
      subString: string,
    [<ExcelArgument(Description="If TRUE or omitted, a and A are considered equal, if FALSE, a and A are different")>]
      ignoreCase: objCell,
    [<ExcelArgument(Description="If TRUE, 'subString' is a regular expression, if FALSE or omitted, 'subString' is a literal")>]
      useRegex: objCell
  ): objCell[,] =
  result {
    let! ic = ignoreCase |> XlObj.toBoolWithDefault true
    let! ur = useRegex |> XlObj.toBoolWithDefault false
    if ur then
      let! r = input |> regexFilter subString ic
      return r
    else
      return
        input
        |> stringFilter (fun s -> if ic then s.ToLowerInvariant().Contains(subString.ToLowerInvariant()) else s.Contains(subString))
  } |> XlObjRange.ofResult
