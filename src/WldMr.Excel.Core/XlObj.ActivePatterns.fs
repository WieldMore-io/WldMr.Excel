namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module ActivePattern =
  let (|ExcelMissingRange|_|) (input: xlObj[,]) =
    if input.GetLength 0 = 1 && input.GetLength 1 = 1 then
      match XlObj.toObj input.[0, 0] with
      | :? ExcelMissing as m -> Some m
      | _ -> None
    else
      None

  let (|ExcelMissing|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? ExcelMissing as m -> Some m
    | _ -> None

  let (|ExcelEmpty|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? ExcelEmpty as m -> Some m
    | _ -> None

  let (|ExcelString|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? string as s -> Some s
    | _ -> None

  let (|ExcelBool|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? bool as s -> Some s
    | _ -> None

  let (|ExcelNum|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? float as s -> Some s
    | _ -> None

  let (|ExcelError|_|) (input: xlObj) =
    match XlObj.toObj input with
    | :? ExcelError as e -> Some e
    | _ -> None

  let (|ExcelNA|_|) (input: xlObj) =
    if input = XlObj.Error.xlNA then
      Some XlObj.Error.xlNA
    else
      None


  module XlObj =
    /// <summary>
    /// True if the input is missing, False otherwise
    /// </summary>
    let isMissing (input: xlObj) =
      match input with
      | ExcelMissing _ -> true
      | _ -> false
