namespace String

open NUnit.Framework

open WldMr.Excel
open WldMr.Excel.Functions


[<TestFixture>]
type ``StringStartsWith``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "Ef", missing, missing)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``handles French accents``() =
    (singleCell "É", "é", missing, missing)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is set``() =
    (singleCell "efg", "E", true |> box |> (~%), missing)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E", false |> box |> (~%), missing)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell false)


[<TestFixture>]
type ``StringStartsWith (Regex)``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "^..g", missing, trueCell)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is set``() =
    (singleCell "efg", "E.G", trueCell, trueCell)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E.G", falseCell, trueCell)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell false)

  [<Test>]
  member __.``only match at start``() =
    (singleCell "eefg", "E.G", trueCell, trueCell)
    |> String.xlStringStartsWith
    |> Array2D.shouldEqual (singleCell false)

[<TestFixture>]
type ``StringEndsWith``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "fG", missing, missing)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``handles French accents``() =
    (singleCell "É", "é", missing, missing)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is false``() =
    (singleCell "efg", "G", trueCell, missing)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "G", falseCell, missing)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell false)


[<TestFixture>]
type ``StringEndWith (Regex)``() =

  [<Test>]
  member __.``arguments are defaulted``() =
    (singleCell "efg", "E..$", missing, trueCell)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell true)

  [<Test>]
  member __.``case sensitive if ignoreCase is false``() =
    (singleCell "efg", "E.G", trueCell, trueCell)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell true)

    (singleCell "efg", "E.G", falseCell, trueCell)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell false)

  [<Test>]
  member __.``only match at end``() =
    (singleCell "eefgg", "E.G", trueCell, trueCell)
    |> String.xlStringEndsWith
    |> Array2D.shouldEqual (singleCell false)
