module WldMr.Excel.Functions.xlRegex

open ExcelDna.Integration
open WldMr.Excel

open System.Text.RegularExpressions


[<ExcelFunction(
    Category= "WldMr Text",
    Description=
      "Returns whether a regex matches a string.\r\n"
      + "(Uses .Net regex engine and syntax)",
    IsThreadSafe= true)>]
let xlRegexMatch
  (
    [<ExcelArgument(Description= "Input text" )>]
      input: xlObj[,],
    [<ExcelArgument(Description= "Regular expression pattern to search for" )>]
      pattern: xlObj[,],
    [<ExcelArgument(
        Description=
          "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, " +
          "ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, " +
          "RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" )>]
      regexOptions: xlObj[,]
  ) : xlObj[,]
  =
  let matchRegex input pattern regexOptions =
    try
      let ms = Regex.Match(input, pattern, regexOptions)
      ms.Success |> XlObj.ofBool |> Ok
    with
      | :? System.ArgumentException as e -> $"Regex error: {e.Message}" |> Error
      | e -> $"{e.Message} ({e.GetType()})" |> Error

  // empty cell would lead to the regex "" which matches everything, not very useful
  let patternTrimmed = pattern |> XlObjRange.trimRange XlObjRange.TrimMode.MissingEmptyStringEmpty
  ArrayFunctionBuilder
    .Add("Input", XlObj.toString, input)
    .Add("Pattern", XlObj.toString, patternTrimmed)
    .Add("Options",
         XlObj.toInt
         |> XlObjParser.map enum<RegexOptions>
         |> XlObjParser.withDefault RegexOptions.None,
         regexOptions)
    .EvalFunction matchRegex
  |> FunctionCall.eval
