namespace global

open FsUnit

type InitMsgUtils() =
  inherit FSharpCustomMessageFormatter()


namespace WldMr.Excel
  open FsUnitTyped
  open WldMr.Excel.Utilities


  module Array2D =
    open System.Diagnostics

    [<DebuggerStepThrough>]
    let shouldEqual (expected: 'a[,]) (actual: 'a[,]) =
      actual |> Array2D.flatten |> shouldEqual (expected |> Array2D.flatten)


  [<AutoOpen>]
  module Range =
    open ExcelDna.Integration
    let singleCell (v:'a) = [[v |> box]] |> array2D
    let emptyArray: obj[,] = [] |> array2D
    let missing = ExcelMissing.Value |> box


  module Test =
    let returnedAnArray (x:obj) =
      x |> should be instanceOfType<obj[,]>
      x :?> obj[,]
