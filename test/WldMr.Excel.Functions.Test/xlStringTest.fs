namespace String


open NUnit.Framework
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Utilities
open WldMr.Excel.String


[<TestFixture>]
type ``StringStartsWith``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "Ef", missing, missing)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``handles French accents``() =
    (singleCell "É", "é", missing, missing)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is set``() =
    (singleCell "efg", "E", true |> box, missing)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E", false |> box, missing)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)


[<TestFixture>]
type ``StringStartsWith (Regex)``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "^..g", missing, true)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is set``() =
    (singleCell "efg", "E.G", true |> box, true)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E.G", false |> box, true)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)

  [<Test>]
  member __.``only match at start``() =
    (singleCell "eefg", "E.G", true |> box, true)
    |> Basic.xlStringStartsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)

[<TestFixture>]
type ``StringEndsWith``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "fG", missing, missing)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``handles French accents``() =
    (singleCell "É", "é", missing, missing)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is false``() =
    (singleCell "efg", "G", true |> box, missing)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "G", false |> box, missing)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)


[<TestFixture>]
type ``StringEndWith (Regex)``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "E..$", missing, true)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is false``() =
    (singleCell "efg", "E.G", true, true)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E.G", false |> box, true)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)

  [<Test>]
  member __.``only match at end``() =
    (singleCell "eefgg", "E.G", true |> box, true)
    |> Basic.xlStringEndsWith
    |> Test.returnedAnArray
    |> Array2D.shouldEqual (singleCell false)
