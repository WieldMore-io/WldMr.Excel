namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module ActivePattern =
  let (|ExcelMissingRange|_|) (input: xlObj[,]) =
    if input.GetLength 0 = 1 && input.GetLength 1 = 1 then
      match (% input.[0, 0]: obj) with
      | :? ExcelMissing as m -> Some m
      | _ -> None
    else
      None

  let (|ExcelMissing|_|) (input: xlObj) =
    match (%input: obj) with
    | :? ExcelMissing as m -> Some m
    | _ -> None

  let (|ExcelEmpty|_|) (input: xlObj) =
    match (%input: obj) with
    | :? ExcelEmpty as m -> Some m
    | _ -> None

  let (|ExcelString|_|) (input: xlObj) =
    match (%input: obj) with
    | :? string as s -> Some s
    | _ -> None

  let (|ExcelBool|_|) (input: xlObj) =
    match (%input: obj) with
    | :? bool as s -> Some s
    | _ -> None

  let (|ExcelNum|_|) (input: xlObj) =
    match (%input: obj) with
    | :? float as s -> Some s
    | _ -> None

  let (|ExcelError|_|) (input: xlObj) =
    match (%input: obj) with
    | :? ExcelError as e -> Some e
    | _ -> None

  let (|ExcelNA|_|) (input: xlObj) =
    if input = XlObj.Error.xlNA then
      Some XlObj.Error.xlNA
    else
      None
