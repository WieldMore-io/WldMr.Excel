namespace Stack

open NUnit.Framework

open WldMr.Excel
open WldMr.Excel.Functions


[<TestFixture>]
type ``xlStackV``() =
  [<Test>]
  member __.``returns empty when inputs are empty``() =
    Stack.xlStackV(emptyArray, emptyArray) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; [2.0 |> XlObj.ofFloat]] |> array2D )

    (
      singleCell 1.0,
      singleCell XlObj.objMissing
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D )

  [<Test>]
  member __.``can stack 1x2 with 2x1``() =
    (
      [[3.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D,
      [[1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]] |> array2D
    )
    |> Stack.xlStackV
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
    Stack.xlStackV(emptyArray, emptyArray) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0
    )
    |> Stack.xlStackH
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]] |> array2D )

    (
      singleCell 1.0,
      singleCell XlObj.objMissing
    )
    |> Stack.xlStackH
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a"
    )
    |> Stack.xlStackH
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat; "a" |> XlObj.ofString]] |> array2D )


[<TestFixture>]
type ``xlTrimNA``() =

  [<Test>]
  member __.``returns empty array when input is NA``() =
    Stack.xlTrimNA(singleCell XlObj.Error.objNA) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``returns empty array when input is empty``() =
    Stack.xlTrimNA(emptyArray) |> Array2D.shouldEqual emptyArray

