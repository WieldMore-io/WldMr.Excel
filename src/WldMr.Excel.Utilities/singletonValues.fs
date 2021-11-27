namespace WldMr.Excel.Utilities

open ExcelDna.Integration


[<AutoOpen>]
module SingletonValue =
  [<RequireQualifiedAccess>]
  module XlObj =
    let objEmpty: obj = ExcelEmpty.Value |> box
    let objMissing = ExcelMissing.Value |> box

    module Error =
      let objNA = ExcelError.ExcelErrorNA |> box
      let objValue = ExcelError.ExcelErrorValue |> box
      let objName = ExcelError.ExcelErrorName |> box
      let objNum = ExcelError.ExcelErrorNum |> box
      let objGettingData = ExcelError.ExcelErrorGettingData |> box
