namespace xlStack

open NUnit.Framework
open FSharpPlus
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Helpers


[<TestFixture>]
type ``xlRangeSubRows``() =
  let emptyArray: obj[,] = [[]] |> array2D

  [<Test>]
  member __.``arguments are defaulted``() =
    let a22 = [[1; 2];[3;4]] |>> (map box) |> array2D
    (a22, ExcelMissing.Value |> box, ExcelMissing.Value |> box)
    |> SubRange.xlRangeSubRows
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a22
