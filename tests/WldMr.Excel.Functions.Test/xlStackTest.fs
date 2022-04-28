namespace Stack

open NUnit.Framework

open WldMr.Excel
open WldMr.Excel.Functions


[<TestFixture>]
type ``xlStackV``() =
  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Stack.xlStackV(emptyArray, emptyArray, missingArray, missingArray, missingArray, missingArray, missingArray)
    |> Array2D.shouldEqual (singleCell XlObj.Error.xlNA)

  [<Test>]
  member __.``works with cells``() =
    (
      singleCell 1.0,
      singleCell 2.0,
      missingArray, missingArray, missingArray, missingArray, missingArray
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; [2.0 |> XlObj.ofFloat]] |> array2D )

    (
      singleCell 1.0,
      singleCell XlObj.xlMissing,
      missingArray, missingArray, missingArray, missingArray, missingArray
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual (singleCell 1.0)

    (
      singleCell 1.0,
      singleCell "a",
      missingArray, missingArray, missingArray, missingArray, missingArray
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual ( [[1.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D )

  [<Test>]
  member __.``can stack 1x2 with 2x1``() =
    (
      [[3.0 |> XlObj.ofFloat]; ["a" |> XlObj.ofString]] |> array2D,
      [[1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]] |> array2D,
      missingArray, missingArray, missingArray, missingArray, missingArray
    )
    |> Stack.xlStackV
    |> Array2D.shouldEqual (
      [[3.0 |> XlObj.ofFloat; XlObj.Error.xlNA]
       ["a" |> XlObj.ofString; XlObj.Error.xlNA]
       [1.0 |> XlObj.ofFloat; 2.0 |> XlObj.ofFloat]
      ]
      |> array2D
    )


[<TestFixture>]
type ``xlStackH``() =
  [<Test>]
  member __.``returns NA when inputs are empty``() =
    Stack.xlStackV(emptyArray, emptyArray, missingArray, missingArray, missingArray, missingArray, missingArray)
    |> Array2D.shouldEqual (singleCell XlObj.Error.xlNA)

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
      singleCell XlObj.xlMissing
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
    Stack.xlTrimNA(singleCell XlObj.Error.xlNA) |> Array2D.shouldEqual emptyArray

  [<Test>]
  member __.``returns empty array when input is empty``() =
    Stack.xlTrimNA(emptyArray) |> Array2D.shouldEqual emptyArray

