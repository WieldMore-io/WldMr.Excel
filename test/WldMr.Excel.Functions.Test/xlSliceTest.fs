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
  member __.``select last row``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |>> (map box) |> array2D
    (a33, -1.0, -1.0, missing, missing)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a33.[2..2, *]

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

  [<Test>]
  member __.``3, 4, 1, -2``() =
    let a56 =
      [
        [11; 12; 13; 14; 15; 16]
        [21; 22; 23; 24; 25; 26]
        [31; 32; 33; 34; 35; 36]
        [41; 42; 43; 44; 45; 46]
        [51; 52; 53; 54; 55; 56]
      ] |>> (map box) |> array2D
    (a56, 3.0, 4.0, 1.0, -2.0)
    |> SubRange.xlSlice
    |> Test.returnedAnArray
    |> Array2D.shouldEqual a56.[2..3, 0..4]
