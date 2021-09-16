namespace WldMr.Excel.Utilities

open ExcelDna.Integration
open System


module Array2D =
  /// <summary>
  /// flattens the array as a list of rows, each row being a list
  /// simplify further processing (although potentially slower)
  /// </summary>
  let flatten array2d =
    [ for x in [ 0 .. (Array2D.length1 array2d) - 1 ] do
      [ for y in [ 0 .. (Array2D.length2 array2d) - 1 ] do
        yield array2d.[x, y] ] ]

[<RequireQualifiedAccess>]
module XlObj =
  let getSize (a: obj[,]): int * int =
    match a.GetLength 0, a.GetLength 1 with
    | 0, _ | _, 0 -> 0, 0
    | 1, 1 when a.[0, 0] = objMissing -> 0, 0
    | ls -> ls

  /// <summary>
  /// Defaults the value if the input is Missing, Empty or ""
  /// </summary>
  let defaultWith defaultFun (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultFun () |> box
    | o -> o

  /// <summary>
  /// Defaults the value if the input is Missing, Empty or ""
  /// </summary>
  let defaultValue v (o: obj) =
    match o with
    | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> v |> box
    | o -> o


  /// Returns a column array from a sequence which elements get boxed
  [<RequireQualifiedAccess>]
  module Column =
    let ofSeqWithEmpty (emptyVal: obj) (r: seq<obj>) =
      let v = r |> Array.ofSeq
      if v.Length = 0 then
        emptyVal
      else
        Array2D.init v.Length 1 (fun i j -> v.[i]) |> box


[<AutoOpen>]
module Error =
  [<RequireQualifiedAccess>]
  module XlObj =
    let errorString (errorMessage: string) = $"#Error! {errorMessage}"

    let ofErrorMessage (errorMessage: string) = $"#Error! {errorMessage}" |> box


[<AutoOpen>]
module ToFunctions =
  [<RequireQualifiedAccess>]
  module XlObj =
    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toInt (o: obj) =
      match o with | :? float as f -> f |> int |> Ok | _ -> "Expected a number" |> Error


    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion, and rejects non-integer float 
    /// (A very small rounding is accepted to address potential rounding issues)
    /// </summary>
    let toIntStrict (o: obj) =
      match o with
      | :? float as f ->
          let error = abs ( f - (f |> int |> float) )
          if error < 1e-8 then
            f |> int |> Ok
          else
            "Expected an integer." |> Error
      | _ -> "Expected an integer." |> Error


    /// <summary>
    /// Tries to extract a float out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toFloat (o: obj) =
      match o with | :? float as f -> f |> Ok | _ -> "Expected a number" |> Error

    /// <summary>
    /// Tries to extract a datetime out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toDate (o: obj) =
      match o with | :? float as f -> f |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

    /// <summary>
    /// Tries to extract a date without time out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toDateNoTime (o: obj) =
      match o with | :? float as f -> f |> int |> float |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

    /// <summary>
    /// try to extract a string without time out of excel cell value
    /// does not attempt any conversion (the original excel value must be a string)
    /// </summary>
    let toString (o: obj) =
      match o with | :? string as s -> s |> Ok | _ -> "Expected a string" |> Error

    /// <summary>
    /// force converts an Excel obj input to a boolean option
    /// - missing, empty, or empty-string values are treated as None (!)
    /// - Errors, exact 0 are treated as False (as expected)
    /// - Strings are treated as False (debatable)
    /// - non-zero numbers are treated as True
    /// behaves similarly to Excel functions that take a range (see =AND(...))
    /// </summary>
    let toBoolOption (o:obj) =
      match o with
      | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> None |> Ok
      | ExcelError _ -> false |> Some |> Ok
      | ExcelNum f -> f = 0.0 |> Some |> Ok
      | ExcelBool b -> Some b |> Ok
      | ExcelString s when s.ToLower() = "true" -> true |> Some |> Ok
      | ExcelString s when s.ToLower() = "false" -> false |> Some |> Ok
      | _ -> Some false |> Ok


    /// <summary>
    /// Force-converts an Excel obj input to a boolean
    /// Behaviour is as close to Excel as possible
    /// - missing, empty are false
    /// - Exactxact 0 are treated as False (as expected)
    /// - Strings are treated as False (debatable)
    /// - non-zero numbers are treated as True
    /// </summary>
    let toBool (o:obj) =
      match o with
      | ExcelEmpty _ | ExcelMissing _ -> false |> Ok
      | ExcelNum f -> f = 0.0 |> Ok
      | ExcelBool b -> b |> Ok
      | ExcelString s when s.ToLower() = "true" -> true |> Ok
      | ExcelString s when s.ToLower() = "false" -> false |> Ok
      | _ -> "Expected a boolean" |> Error


    /// <summary>
    /// Force-converts an Excel obj input to a boolean, using a default value if missing or empty
    /// Behaviour is as close to Excel as possible
    /// - Exact 0 are treated as False (as expected)
    /// - Strings are treated as False (debatable)
    /// - non-zero numbers are treated as True
    /// </summary>
    let toBoolWithDefault d (o:obj) =
      match o with
      | ExcelMissing _ -> d |> Ok
      | _ -> o |> toBool


[<AutoOpen>]
module OfFunctions =
  [<RequireQualifiedAccess>]
  module XlObj =
    /// <summary>
    /// boxes the 
    /// </summary>
    let ofString (s: string) =
      s |> box

    /// <summary>
    /// </summary>
    let ofArray2dWithEmpty (emptyValue: obj) (a: obj[,]) =
      if a.Length = 0 then
        emptyValue
      else
        a |> box


    /// <summary>
    /// Boxes the float
    /// If its value is NaN, it is replaced by #N/A!.
    /// does not attempt any conversion (the original excel value must be a string)
    /// </summary>
    let ofFloat f =
      if Double.IsNaN f then
        objNA
      else
        f |> box

    /// <summary>
    /// Summarizes the number of errors and then list them
    /// </summary>
    let ofValidation (t: Result<obj, string list>): obj =
      let errorMessage errors =
        let sep = "; "
        match errors with
        | [] -> "Unexpected error"
        | x::[] -> x
        | xs -> $"{xs.Length} errors: {String.Join(sep, xs)}"

      match t with
      | Ok v -> v
      | Error e -> e |> errorMessage |> XlObj.ofErrorMessage

    /// <summary>
    /// Converts a Result of obj into a suitable valid Excel output value
    /// </summary>
    let ofResult (t: Result<obj, string>): obj =
      match t with
      | Ok v -> v
      | Error err -> err |> XlObj.ofErrorMessage


[<AutoOpen>]
module ArgToFunctions =
  [<RequireQualifiedAccess>]
  module XlObj =
    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion or rounding (the original excel value must be a number)
    /// (A very small rounding is attempted)
    /// </summary>
    let argToIntStrict (argName: string) (o: obj) =
      match o with
      | :? float as f ->
          let error = abs ( f - (f |> int |> float) )
          if error < 1e-8 then
            f |> int |> Ok
          else
            $"Argument '{argName}': expected an integer." |> Error
      | _ -> $"Argument '{argName}': expected an integer." |> Error

    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let argToInt (argName: string) (o: obj) =
      match o with
      | :? float as f -> f |> int |> Ok
      | _ -> $"Argument '{argName}': expected a number." |> Error

    /// <summary>
    /// Tries to extract a float out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let argToFloat (argName: string) (o: obj) =
      match o with
      | :? float as f -> f |> Ok
      | _ -> $"Argument '{argName}': expected a number." |> Error

    /// <summary>
    /// try to extract a string without time out of excel cell value
    /// does not attempt any conversion (the original excel value must be a string)
    /// </summary>
    let argToString (argName: string) (o: obj) =
      match o with
      | :? string as s -> s |> Ok
      | _ -> $"Argument '{argName}': expected a string." |> Error


module Result =
  let mapArgError (errMsg: string) = Result.mapError (fun e -> [$"Arg '{errMsg}': {e}"])
