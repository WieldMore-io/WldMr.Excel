namespace Stack

open NUnit.Framework
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Utilities


[<TestFixture>]
type ``xlStackV``() =
  [<Test>]
  member __.``returns empty when inputs are empty``() =
    Range.xlStackV(emptyArray, emptyArray) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Range.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; [2.0 |> XlObj.ofFloat]] |> array2D )

    (
      singleCell 1.0,
      singleCell ExcelMissing.Value
    )
    |> Range.xlStackV
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Range.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D )

  [<Test>]
  member __.``can stack 1x2 with 2x1``() =
    (
      [[3.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D,
      [[1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]] |> array2D
    )
    |> Range.xlStackV
    |> Array2D.shouldEqual (
      [[3.0 |> XlObj.ofFloat; XlObj.Error.objNA]
       ["a" |> XlObj.ofString; XlObj.Error.objNA]
       [1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]
      ]
      |> array2D
    )


[<TestFixture>]
type ``xlStackH``() =
  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Range.xlStackV(emptyArray, emptyArray) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Range.xlStackH
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]] |> array2D )

    (
      singleCell 1.0,
      singleCell ExcelMissing.Value
    )
    |> Range.xlStackH
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Range.xlStackH
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat; "a" |> XlObj.ofString]] |> array2D )


[<TestFixture>]
type ``xlTrimNA``() =

  [<Test>]
  member __.``returns empty array when input is NA``() =
    Range.xlTrimNA(singleCell ExcelError.ExcelErrorNA) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``returns empty array when input is empty``() =
    Range.xlTrimNA(emptyArray) |> Array2D.shouldEqual emptyArray

