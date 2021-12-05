namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module SingletonValue =
  [<RequireQualifiedAccess>]
  module XlObj =
    let xlEmpty: objCell = ExcelEmpty.Value |> box |> (~%)
    let xlMissing: objCell = ExcelMissing.Value |> box |> (~%)

    module Error =
      let xlNA: objCell = ExcelError.ExcelErrorNA |> box |> (~%)
      let xlValue: objCell = ExcelError.ExcelErrorValue |> box |> (~%)
      let xlName: objCell = ExcelError.ExcelErrorName |> box |> (~%)
      let xlNum: objCell = ExcelError.ExcelErrorNum |> box |> (~%)
      let xlGettingData: objCell = ExcelError.ExcelErrorGettingData |> box |> (~%)
      let xlDiv0: objCell = ExcelError.ExcelErrorDiv0 |> box |> (~%)
      let xlNull: objCell = ExcelError.ExcelErrorNull |> box |> (~%)
      let xlRef: objCell = ExcelError.ExcelErrorRef |> box |> (~%)
