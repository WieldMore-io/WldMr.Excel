namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module ActivePattern =
  let (|ExcelMissingRange|_|) (input: objCell[,]) =
    if input.GetLength 0 = 1 && input.GetLength 1 = 1 then
      match (% input.[0, 0]: obj) with
      | :? ExcelMissing as m -> Some m
      | _ -> None
    else
      None

  let (|ExcelMissing|_|) (input: objCell) =
    match (%input: obj) with
    | :? ExcelMissing as m -> Some m
    | _ -> None

  let (|ExcelEmpty|_|) (input: objCell) =
    match (%input: obj) with
    | :? ExcelEmpty as m -> Some m
    | _ -> None

  let (|ExcelString|_|) (input: objCell) =
    match (%input: obj) with
    | :? string as s -> Some s
    | _ -> None

  let (|ExcelBool|_|) (input: objCell) =
    match (%input: obj) with
    | :? bool as s -> Some s
    | _ -> None

  let (|ExcelNum|_|) (input: objCell) =
    match (%input: obj) with
    | :? float as s -> Some s
    | _ -> None

  let (|ExcelError|_|) (input: objCell) =
    match (%input: obj) with
    | :? ExcelError as e -> Some e
    | _ -> None

  let (|ExcelNA|_|) (input: objCell) =
    if input = XlObj.Error.objNA then
      Some XlObj.Error.objNA
    else
      None
