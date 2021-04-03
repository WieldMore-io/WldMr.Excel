module WldMr.Excel.String.Basic

open ExcelDna.Integration
open FSharpPlus
open WldMr.Excel.Helpers


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
      prefix:string
  ): obj[,] =
  input
  |> stringFilter (String.startsWith prefix)


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
      suffix:string
  ): obj[,] =
  input
  |> stringFilter (String.endsWith suffix)


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
      subString:string
  ): obj[,] =
  input
  |> stringFilter (String.isSubString subString)
