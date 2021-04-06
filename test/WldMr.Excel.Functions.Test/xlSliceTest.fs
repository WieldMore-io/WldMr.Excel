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
    let a22 = [[1; 2]; [3;4]] |>> (map box) |> array2D
    (a22, missing, missing, missing, missing)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a22

  [<Test>]
  member __.``select rows``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |>> (map box) |> array2D
    (a33, 2.0, -2.0, missing, missing)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a33.[1..1, *]

  [<Test>]
  member __.``select columns``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |>> (map box) |> array2D
    (a33, 2.0, -2.0, missing, missing)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a33.[1..1, *]

  [<Test>]
  member __.``select center``() =
    let a33 = [[11; 12; 13]; [21;22;23]; [31; 32; 33]] |>> (map box) |> array2D
    (a33, 2.0, -2.0, 2.0, -2.0)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a33.[1..1, 1..1]
