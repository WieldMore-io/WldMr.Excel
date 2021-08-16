namespace WldMr.Excel.Utilities

open ExcelDna.Integration
open System


module Array2D =
  /// flattens the array as a list of rows, each row being a list
  /// simplify further processing (although potentially slower)
  let flatten array2d =
    [ for x in [ 0 .. (Array2D.length1 array2d) - 1 ] do
      [ for y in [ 0 .. (Array2D.length2 array2d) - 1 ] do
        yield array2d.[x, y] ] ]


module XlObj =
  let getSize (a: obj[,]): int * int =
    match a.GetLength 0, a.GetLength 1 with
    | 0, _ | _, 0 -> 0, 0
    | 1, 1 when a.[0, 0] = objMissing -> 0, 0
    | ls -> ls

  /// Defaults the value if the input is Missing, Empty or ""
  let defaultWith defaultFun (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultFun () |> box
    | o -> o

  let defaultValue v (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> v |> box
    | o -> o

  /// Tries to extract an int out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toInt (o: obj) =
    match o with | :? float as f -> f |> int |> Ok | _ -> "Expected a number" |> Error

  /// Tries to extract a float out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toFloat (o: obj) =
    match o with | :? float as f -> f |> Ok | _ -> "Expected a number" |> Error

  /// Tries to extract a datetime out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toDate (o: obj) =
    match o with | :? float as f -> f |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  /// Tries to extract a date without time out of excel cell value
  /// Does not attempt any conversion (the original excel value must be a number)
  let toDateNoTime (o: obj) =
    match o with | :? float as f -> f |> int |> float |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

  /// try to extract a string without time out of excel cell value
  /// does not attempt any conversion (the original excel value must be a string)
  let toString (o: obj) =
    match o with | :? string as s -> s |> Ok | _ -> "Expected a string" |> Error

  /// force converts an Excel obj input to a boolean option
  /// - missing, empty, or empty-string values are treated as None (!)
  /// - Errors, exact 0 are treated as False (as expected)
  /// - Strings are treated as False (debatable)
  /// - non-zero numbers are treated as True
  /// behaves similarly to Excel functions that take a range (see =AND(...))
  let toBoolOption (o:obj) =
    match o with
    | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> None |> Ok
    | ExcelError _ -> false |> Some |> Ok
    | ExcelNum f -> f = 0.0 |> Some |> Ok
    | ExcelBool b -> Some b |> Ok
    | ExcelString s when s.ToLower() = "true" -> true |> Some |> Ok
    | ExcelString s when s.ToLower() = "false" -> false |> Some |> Ok
    | _ -> Some false |> Ok


  /// Force-converts an Excel obj input to a boolean
  /// Behaviour is as close to Excel as possible
  /// - missing, empty are false
  /// - Exactxact 0 are treated as False (as expected)
  /// - Strings are treated as False (debatable)
  /// - non-zero numbers are treated as True
  let toBool (o:obj) =
    match o with
    | ExcelEmpty _ | ExcelMissing _ -> false |> Ok
    | ExcelNum f -> f = 0.0 |> Ok
    | ExcelBool b -> b |> Ok
    | ExcelString s when s.ToLower() = "true" -> true |> Ok
    | ExcelString s when s.ToLower() = "false" -> false |> Ok
    | _ -> "Expected a boolean" |> Error


  /// Force-converts an Excel obj input to a boolean, using a default value if missing or empty
  /// Behaviour is as close to Excel as possible
  /// - Exact 0 are treated as False (as expected)
  /// - Strings are treated as False (debatable)
  /// - non-zero numbers are treated as True
  let toBoolWithDefault d (o:obj) =
    match o with
    | ExcelMissing _ -> d |> Ok
    | _ -> o |> toBool


  /// Boxes the float
  /// If its value is NaN, it is replaced by #N/A!.
  /// does not attempt any conversion (the original excel value must be a string)
  let ofFloat f =
    if Double.IsNaN f then
      objNA
    else
      f |> box

  /// Summarizes the number of errors and returns them prefixed by "#Error!""
  let ofValidation (t: Result<obj, string list>): obj =
    let errorMessage errors =
      let sep = "; "
      match errors with
      | [] -> "Unexpected error"
      | x::[] -> $"#Error! {x}"
      | xs -> $"#Error! {xs.Length} errors: {String.Join(sep, xs)}"

    match t with
    | Ok v -> v
    | Error e -> e |> errorMessage |> box

  let ofResult<'E> (t: Result<obj, 'E>): obj =
    match t with
    | Ok v -> v
    | Error err -> $"#Error! {err}" :> obj


  /// Returns a column array from a sequence which elements get boxed
  let inline columnOfSeq ifEmpty r =
    let v = r |> Seq.map box |> Array.ofSeq
    let a =
      if v.Length = 0 then
        Array2D.create 1 1 (ifEmpty |> box)
      else
        Array2D.init v.Length 1 (fun i j -> v.[i])
    a |> box


module Result =
  let mapArgError errMsg = Result.mapError (fun e -> [$"Arg '{errMsg}': {e}"])
