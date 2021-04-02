namespace xlStack

open NUnit.Framework
open FSharpPlus
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Helpers


[<TestFixture>]
type ``xlStackV``() =
  let emptyArray: obj[,] = [[]] |> array2D

  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Range.xlStackV(emptyArray, emptyArray) |> should be (equal ExcelError.ExcelErrorNA)

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Range.xlStackV
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box]; [2.0 |> box]] |> array2D )

    (
      singleCell 1.0,
      singleCell ExcelMissing.Value
    )
    |> Range.xlStackV
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Range.xlStackV
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box]; ["a" |> box]] |> array2D )


[<TestFixture>]
type ``xlStackH``() =

  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Range.xlStackV(emptyArray, emptyArray) |> should be (equal ExcelError.ExcelErrorNA)

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Range.xlStackH
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box; 2.0 |> box]] |> array2D )

    (
      singleCell 1.0,
      singleCell ExcelMissing.Value
    )
    |> Range.xlStackH
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Range.xlStackH
    |> (fun x -> x :?> obj[,])
    |> Array2D.shouldEqual ( [[1.0 |> box; "a" |> box]] |> array2D )


[<TestFixture>]
type ``xlTrimNA``() =

  [<Test>]
  member __.``returns empty array when input is NA``() =
    Range.xlTrimNA(singleCell ExcelError.ExcelErrorNA) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``returns empty array when input is empty``() =
    Range.xlTrimNA(emptyArray) |> Array2D.shouldEqual emptyArray

