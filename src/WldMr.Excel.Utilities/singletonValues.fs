namespace WldMr.Excel.Utilities

open ExcelDna.Integration


[<AutoOpen>]
module SingletonValue =
  [<RequireQualifiedAccess>]
  module XlObj =
    let objEmpty: obj = ExcelEmpty.Value |> box
    let objMissing = ExcelMissing.Value |> box
    let objNA = ExcelError.ExcelErrorNA |> box
    let objName = ExcelError.ExcelErrorName |> box
    let objGettingData = ExcelError.ExcelErrorGettingData |> box
