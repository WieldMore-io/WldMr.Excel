namespace global

open FsUnit

type InitMsgUtils() =
  inherit FSharpCustomMessageFormatter()


namespace WldMr.Excel
  open FsUnitTyped
  open WldMr.Excel


  module Array2D =
    open System.Diagnostics

    [<DebuggerStepThrough>]
    let shouldEqual (expected: 'a[,]) (actual: 'a[,]) =
      actual |> Array2D.flatten |> shouldEqual (expected |> Array2D.flatten)


  [<AutoOpen>]
  module Range =
    let singleCell (v:'a): objCell[,] = [[v |> box |> (~%)]] |> array2D
    let emptyArray: objCell[,] = [] |> array2D
    let trueCell: objCell = true |> XlObj.ofBool
    let falseCell: objCell = false |> XlObj.ofBool
    let missing: objCell = XlObj.objMissing
