namespace Range

open NUnit.Framework
open FSharpPlus
open FsUnit
open FsUnitTyped

open WldMr.Excel
open ExcelDna.Integration
open WldMr.Excel.Helpers


[<TestFixture>]
type ``xlSlice``() =
  let emptyArray: obj[,] = [[]] |> array2D

  [<Test>]
  member __.``arguments are defaulted``() =
    let a22 = [[1; 2];[3;4]] |>> (map box) |> array2D
    (a22, missing, missing, missing, missing)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a22
