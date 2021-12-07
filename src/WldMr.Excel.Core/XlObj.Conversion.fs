namespace WldMr.Excel

open System


[<RequireQualifiedAccess>]
module XlObj =

  /// <summary>
  /// True if the value is missing, False otherwise
  /// </summary>
  let isMissing (o: xlObj): bool =
    match o with
    | ExcelMissing _ -> true
    | _ -> false


[<AutoOpen>]
module Error =
  [<RequireQualifiedAccess>]
  module XlObj =
    let errorString (errorMessage: string) = $"#Error! {errorMessage}"

    let ofErrorMessage (errorMessage: string): xlObj = $"#Error! {errorMessage}" |> box |> (~%)


[<AutoOpen>]
module ToFunctions =
  [<RequireQualifiedAccess>]
  module XlObj =
    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toInt (o: xlObj) =
      match o with | ExcelNum f -> f |> int |> Ok | _ -> "Expected a number" |> Error

    let toIntDefault defaultValue (o: xlObj) =
      match o with
      | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultValue |> Ok
      | ExcelNum f -> f |> int |> Ok
      | _ -> "Expected a number" |> Error


    /// <summary>
    /// Tries to extract an int out of excel cell value
    /// Does not attempt any conversion, and rejects non-integer float
    /// (A very small rounding is accepted to address potential rounding issues)
    /// </summary>
    let toIntStrict (o: xlObj) =
      match o with
      | ExcelNum f ->
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
    let toFloat (o: xlObj) =
      match o with | ExcelNum f -> f |> Ok | _ -> "Expected a number" |> Error


    let toFloatDefault defaultValue (o: xlObj) =
      match o with
      | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultValue
      | ExcelNum f -> f |> Ok
      | _ -> "Expected a number" |> Error

    /// <summary>
    /// Tries to extract a datetime out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toDate (o: xlObj) =
      match o with | ExcelNum f -> f |> DateTime.FromOADate |> Ok | _ -> "Expected a date" |> Error

    let toDateDefault defaultValue (o: xlObj) =
      match o with
      | ExcelMissing _ | ExcelEmpty _ | ExcelString "" -> defaultValue |> Ok
      | ExcelNum f -> f |> DateTime.FromOADate |> Ok
      | _ -> "Expected a date" |> Error


    /// <summary>
    /// Tries to extract a date without time out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let toDateNoTime (o: xlObj) =
      match o with
      | ExcelNum f -> f |> int |> float |> DateTime.FromOADate |> Ok
      | _ -> "Expected a date" |> Error

    /// <summary>
    /// try to extract a string without time out of excel cell value
    /// does not attempt any conversion (the original excel value must be a string)
    /// </summary>
    let toString (o: xlObj) =
      match o with
      | ExcelString s -> s |> Ok
      | _ -> "Expected a string" |> Error

    /// <summary>
    /// force converts an Excel obj input to a boolean option
    /// - missing, empty, or empty-string values are treated as None (!)
    /// - Errors, exact 0 are treated as False (as expected)
    /// - Strings are treated as False (debatable)
    /// - non-zero numbers are treated as True
    /// behaves similarly to Excel functions that take a range (see =AND(...))
    /// </summary>
    let toBoolOption (o:xlObj) =
      match o with
      | ExcelEmpty _ | ExcelMissing _ | ExcelString "" -> None |> Ok
      | ExcelError _ -> false |> Some |> Ok
      | ExcelNum f -> f = 0.0 |> Some |> Ok
      | ExcelBool b -> Some b |> Ok
      | ExcelString s when s.ToLower() = "true" -> true |> Some |> Ok
      | ExcelString s when s.ToLower() = "false" -> false |> Some |> Ok
      | ExcelString s -> $"Expected a boolean '{s}'" |> Error
      | _ -> Some false |> Ok


    /// <summary>
    /// Force-converts an Excel obj input to a boolean
    /// Behaviour is as close to Excel as possible
    /// - missing, empty are false
    /// - Exact 0 are treated as False (as expected)
    /// - Strings are treated as False (debatable)
    /// - non-zero numbers are treated as True
    /// </summary>
    let toBool (o:xlObj) =
      match o with
      | ExcelEmpty _ | ExcelMissing _ -> false |> Ok
      | ExcelNum f -> f <> 0.0 |> Ok
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
    let toBoolWithDefault d (o:xlObj) =
      match o with
      | ExcelMissing _ -> d |> Ok
      | _ -> o |> toBool


[<AutoOpen>]
module OfFunctions =
  [<RequireQualifiedAccess>]
  module XlObj =
    /// <summary>
    /// boxes a boolean
    /// </summary>
    let ofBool (b: bool): xlObj =
      b |> box |> (~%)

    /// <summary>
    /// boxes a string
    /// </summary>
    let ofString (s: string): xlObj =
      s |> box |> (~%)

    /// <summary>
    /// Boxes the float
    /// If its value is NaN, it is replaced by #N/A!.
    /// </summary>
    let ofFloat f: xlObj  =
      if Double.IsNaN f then
        XlObj.Error.xlNA
      else
        f |> box |> (~%)

    /// <summary>
    /// Boxes the int
    /// </summary>
    let ofInt i: xlObj  =
      i |> float |> box |> (~%)


    let ofDate (d: DateTime): xlObj  =
      d |> box |> (~%)


    /// <summary>
    /// Summarizes the number of errors and then list them
    /// </summary>
    let ofValidation (t: Result<xlObj, string list>): xlObj =
      let errorMessage errors =
        let sep = "; "
        match errors with
        | [] -> "Unexpected error"
        | [ x ] -> x
        | xs -> $"{xs.Length} errors: {String.Join(sep, xs)}"

      match t with
      | Ok v -> v
      | Error e -> e |> errorMessage |> XlObj.ofErrorMessage

    /// <summary>
    /// Converts a Result of obj into a suitable valid Excel output value
    /// </summary>
    let ofResult (t: Result<xlObj, string>): xlObj =
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
    let argToIntStrict (argName: string) (o: xlObj) =
      match o with
      | ExcelNum f ->
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
    let argToInt (argName: string) (o: xlObj) =
      match o with
      | ExcelNum f -> f |> int |> Ok
      | _ -> $"Argument '{argName}': expected a number." |> Error

    /// <summary>
    /// Tries to extract a float out of excel cell value
    /// Does not attempt any conversion (the original excel value must be a number)
    /// </summary>
    let argToFloat (argName: string) (o: xlObj) =
      match o with
      | ExcelNum f -> f |> Ok
      | _ -> $"Argument '{argName}': expected a number." |> Error

    /// <summary>
    /// try to extract a string without time out of excel cell value
    /// does not attempt any conversion (the original excel value must be a string)
    /// </summary>
    let argToString (argName: string) (o: xlObj) =
      match o with
      | ExcelString s -> s |> Ok
      | _ -> $"Argument '{argName}': expected a string." |> Error

    let argToDate (argName: string) (o: xlObj) =
      match o with
      | ExcelNum f -> f |> DateTime.FromOADate |> Ok
      | _ -> $"Argument '{argName}': expected a date." |> Error


    let argDefault defaultValue (argParse: _ -> xlObj -> _) name (o: xlObj) =
      match o with
      | ExcelMissing _ | ExcelEmpty _ -> defaultValue |> Ok
      | _ -> argParse name o
