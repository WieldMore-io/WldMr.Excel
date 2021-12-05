namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module SingletonValue =
  [<RequireQualifiedAccess>]
  module XlObj =
    let objEmpty: objCell = ExcelEmpty.Value |> box |> (~%)
    let objMissing: objCell = ExcelMissing.Value |> box |> (~%)

    module Error =
      let objNA: objCell = ExcelError.ExcelErrorNA |> box |> (~%)
      let objValue: objCell = ExcelError.ExcelErrorValue |> box |> (~%)
      let objName: objCell = ExcelError.ExcelErrorName |> box |> (~%)
      let objNum: objCell = ExcelError.ExcelErrorNum |> box |> (~%)
      let objGettingData: objCell = ExcelError.ExcelErrorGettingData |> box |> (~%)
      let objDiv0: objCell = ExcelError.ExcelErrorDiv0 |> box |> (~%)
      let objNull: objCell = ExcelError.ExcelErrorNull |> box |> (~%)
      let objRef: objCell = ExcelError.ExcelErrorRef |> box |> (~%)
