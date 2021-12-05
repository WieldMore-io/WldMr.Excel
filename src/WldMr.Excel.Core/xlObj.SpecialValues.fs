namespace WldMr.Excel

open ExcelDna.Integration


[<AutoOpen>]
module SingletonValue =
  [<RequireQualifiedAccess>]
  module XlObj =
    let xlEmpty: xlObj = ExcelEmpty.Value |> box |> (~%)
    let xlMissing: xlObj = ExcelMissing.Value |> box |> (~%)

    module Error =
      let xlNA: xlObj = ExcelError.ExcelErrorNA |> box |> (~%)
      let xlValue: xlObj = ExcelError.ExcelErrorValue |> box |> (~%)
      let xlName: xlObj = ExcelError.ExcelErrorName |> box |> (~%)
      let xlNum: xlObj = ExcelError.ExcelErrorNum |> box |> (~%)
      let xlGettingData: xlObj = ExcelError.ExcelErrorGettingData |> box |> (~%)
      let xlDiv0: xlObj = ExcelError.ExcelErrorDiv0 |> box |> (~%)
      let xlNull: xlObj = ExcelError.ExcelErrorNull |> box |> (~%)
      let xlRef: xlObj = ExcelError.ExcelErrorRef |> box |> (~%)
