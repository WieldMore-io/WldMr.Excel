namespace xlStack

open NUnit.Framework
open FSharpPlus
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Helpers



module Array2D =
  open System.Diagnostics

  [<DebuggerStepThrough>]
  let shouldEqual (actual: 'a[,]) (expected: 'a[,]) =
    actual |> Array2D.flattenArray |> shouldEqual (expected |> Array2D.flattenArray)


[<AutoOpen>]
module Range =
  let singleCell (v:'a) = [[v |> box]] |> array2D


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


[<TestFixture>]
type ``xlStackH``() =
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
    |> Range.xlStackH
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box; 2.0 |> box]] |> array2D )


    ([[1.0 |> box]] |> array2D, [[ExcelMissing.Value |> box]] |> array2D)
    |> Range.xlStackH
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box]] |> array2D )



[<TestFixture>]
type ``xlTrimNA``() =
  let emptyArray: obj[,] = [[]] |> array2D

  [<Test>]
  member __.``returns empty array when input is NA``() =
    Range.xlTrimNA(singleCell ExcelError.ExcelErrorNA) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``returns empty array when input is empty``() =
    Range.xlTrimNA(emptyArray) |> Array2D.shouldEqual emptyArray

