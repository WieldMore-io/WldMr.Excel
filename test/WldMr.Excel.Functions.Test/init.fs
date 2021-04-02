namespace global

open FsUnit

type InitMsgUtils() =
  inherit FSharpCustomMessageFormatter()


namespace WldMr.Excel
  open FsUnitTyped
  open WldMr.Excel.Helpers


  module Array2D =
    open System.Diagnostics

    [<DebuggerStepThrough>]
    let shouldEqual (actual: 'a[,]) (expected: 'a[,]) =
      actual |> Array2D.flattenArray |> shouldEqual (expected |> Array2D.flattenArray)


  [<AutoOpen>]
  module Range =
    let singleCell (v:'a) = [[v |> box]] |> array2D
    let emptyArray: obj[,] = [[]] |> array2D


  module Test =
    let returnedAnArray (x:obj) =
      x |> should be instanceOfType<obj[,]>
      x :?> obj[,]
