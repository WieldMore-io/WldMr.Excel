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
  Description= """ """,
  HelpTopic=""
)
  >]
let xlStringStartsWith (input: obj[,], subString:string): obj[,] =
  input
  |> stringFilter (String.startsWith subString)


[<ExcelFunction(Category= "WldMr.String",
  Description= """ """,
  HelpTopic=""
)
>]
let xlStringEndsWith (input: obj[,], subString:string): obj[,] =
  input
  |> stringFilter (String.endsWith subString)

[<ExcelFunction(Category= "WldMr.String",
  Description= """ """,
  HelpTopic=""
)
>]
let xlStringContains (input: obj[,], subString:string): obj[,] =
  input
  |> stringFilter (String.isSubString subString)
