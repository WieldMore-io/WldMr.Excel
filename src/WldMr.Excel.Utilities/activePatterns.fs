namespace WldMr.Excel.Utilities

open ExcelDna.Integration


[<AutoOpen>]
module ActivePattern =
  let (|ExcelMissingRange|_|) (input: obj[,]) =
    if input.GetLength 0 = 1 && input.GetLength 1 = 1 then
      match input.[0, 0] with
      | :? ExcelMissing as m -> Some m
      | _ -> None
    else
      None

  let (|ExcelMissing|_|) (input: obj) =
    match input with
    | :? ExcelMissing as m -> Some m
    | _ -> None

  let (|ExcelEmpty|_|) (input: obj) =
    match input with
    | :? ExcelEmpty as m -> Some m
    | _ -> None

  let (|ExcelString|_|) (input: obj) =
    match input with
    | :? string as s -> Some s
    | _ -> None

  let (|ExcelError|_|) (input: obj) =
    match input with
    | :? ExcelError as e -> Some e
    | _ -> None

  let (|ExcelBool|_|) (input: obj) =
    match input with
    | :? bool as s -> Some s
    | _ -> None

  let (|ExcelNum|_|) (input: obj) =
    match input with
    | :? float as s -> Some s
    | _ -> None
