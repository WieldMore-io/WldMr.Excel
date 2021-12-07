namespace WldMr.Excel

open ExcelDna.Integration


[<RequireQualifiedAccess>]
module XlObj =
  let xlEmpty: xlObj = ExcelEmpty.Value |> XlObj.Unsafe.ofObj
  let xlMissing: xlObj = ExcelMissing.Value |> XlObj.Unsafe.ofObj

  [<RequireQualifiedAccess>]
  module Error =
    let xlNA: xlObj = ExcelError.ExcelErrorNA |> XlObj.Unsafe.ofObj
    let xlValue: xlObj = ExcelError.ExcelErrorValue |> XlObj.Unsafe.ofObj
    let xlName: xlObj = ExcelError.ExcelErrorName |> XlObj.Unsafe.ofObj
    let xlNum: xlObj = ExcelError.ExcelErrorNum |> XlObj.Unsafe.ofObj
    let xlGettingData: xlObj = ExcelError.ExcelErrorGettingData |> XlObj.Unsafe.ofObj
    let xlDiv0: xlObj = ExcelError.ExcelErrorDiv0 |> XlObj.Unsafe.ofObj
    let xlNull: xlObj = ExcelError.ExcelErrorNull |> XlObj.Unsafe.ofObj
    let xlRef: xlObj = ExcelError.ExcelErrorRef |> XlObj.Unsafe.ofObj

  /// <summary>
  /// Boxes the boolean into an xlObj phantom type
  /// </summary>
  let ofBool (b: bool): xlObj =
    b |> XlObj.Unsafe.ofObj

  /// <summary>
  /// Boxes the string into an xlObj phantom type
  /// </summary>
  let ofString (s: string): xlObj =
    s |> XlObj.Unsafe.ofObj

  /// <summary>
  /// Boxes the float into an xlObj phantom type
  /// If its value is NaN, it is replaced by #N/A!.
  /// </summary>
  let ofFloat f: xlObj  =
    if System.Double.IsNaN f then
      Error.xlNA
    else
      f |> XlObj.Unsafe.ofObj
