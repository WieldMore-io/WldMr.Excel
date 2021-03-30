namespace xlStack

open NUnit.Framework
open FSharpPlus
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Helpers



[<AutoOpen>]
module Array2D =
  open System.Diagnostics
  open NUnit.Framework
  open System.Collections.Generic

  [<DebuggerStepThrough>]
  let shouldEqual (expected: 'a[,]) (actual: 'a[,]) =
    actual |> Array2D.flattenArray |> shouldEqual (expected |> Array2D.flattenArray)


[<TestFixture>]
type ``xlStackV``() =
  let emptyArray: obj[,] = [[]] |> array2D

  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Range.xlStackV(emptyArray, emptyArray) |> should be (equal ExcelError.ExcelErrorNA)
          
  [<Test>]
  member __.``works with cells``() =
    (
      [[1.0 |> box]] |> array2D,
      [[2.0 |> box]] |> array2D
    )
    |> Range.xlStackV
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box]; [2.0 |> box]] |> array2D )


    ([[1.0 |> box]] |> array2D, [[ExcelMissing.Value |> box]] |> array2D)
    |> Range.xlStackV
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box]] |> array2D )

