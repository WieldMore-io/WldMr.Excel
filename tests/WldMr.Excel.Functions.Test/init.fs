namespace global

open FsUnit

type InitMsgUtils() =
  inherit FSharpCustomMessageFormatter()


namespace WldMr.Excel
  open FsUnitTyped
  open WldMr.Excel
  open WldMr.Excel.Core.Extensions


  module Array2D =
    open System.Diagnostics

    [<DebuggerStepThrough>]
    let shouldEqual (expected: 'a[,]) (actual: 'a[,]) =
      actual |> Array2D.flatten |> shouldEqual (expected |> Array2D.flatten)


  [<AutoOpen>]
  module Range =
    let singleCell (v:'a): xlObj[,] = [[v |> box |> (~%)]] |> array2D
    let emptyArray: xlObj[,] = [] |> array2D
    let trueCell: xlObj = true |> XlObj.ofBool
    let falseCell: xlObj = false |> XlObj.ofBool
    let missing: xlObj = XlObj.xlMissing
