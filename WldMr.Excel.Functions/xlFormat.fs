module WldMr.Excel.String.Format

open ExcelDna.Integration
open FSharpPlus

open WldMr.Excel.Helpers

[<ExcelFunction(Category= "WldMr.String", 
  Description= """Format an interpolated string
eg xlFormat("dd-mmm-yy", A1, "d")
https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types""",
  HelpTopic="https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types"
)
  >]
let xlFormatA (s:string, o1:obj, t1:obj, o2: obj, t2:obj, o3: obj, t3:obj, o4: obj, t4:obj) =
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

  let resToFlatten = monad {
    let! a1 = formatValue o1 t1
    and! a2 = formatValue o2 t2
    and! a3 = formatValue o3 t3
    and! a4 = formatValue o4 t4
    Result.protect (fun () -> System.String.Format(s, a1, a2, a3, a4) :> obj) () |> Result.mapError (fun err -> $"{err}")
  }
  let res = resToFlatten |> Result.flatten
  res |> XlObj.ofResult
