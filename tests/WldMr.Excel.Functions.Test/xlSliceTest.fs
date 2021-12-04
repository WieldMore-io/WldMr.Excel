namespace Range

open NUnit.Framework

open WldMr.Excel
open WldMr.Excel.Functions


[<TestFixture>]
type ``xlSlice``() =

  let toArrayObj a: objCell[,] = a |> List.map (List.map (box >> (~%))) |> array2D

  [<Test>]
  member __.``arguments are defaulted``() =
    let a22 = [[1; 2]; [3;4]] |> toArrayObj
    (a22, missing, missing, missing, missing)
    |> Slice.xlSlice
    |> Array2D.shouldEqual a22

  [<Test>]
  member __.``select rows``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |> toArrayObj
    (a33, 2.0 |> XlObj.ofFloat, -2.0 |> XlObj.ofFloat, missing, missing)
    |> Slice.xlSlice
    |> Array2D.shouldEqual a33.[1..1, *]

  [<Test>]
  member __.``select last row``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |> toArrayObj
    (a33, -1.0 |> XlObj.ofFloat, -1.0 |> XlObj.ofFloat, missing, missing)
    |> Slice.xlSlice
    |> Array2D.shouldEqual a33.[2..2, *]

  [<Test>]
  member __.``select columns``() =
    let a33 = [[11; 12; 13];[21;22;23];[31; 32; 33]] |> toArrayObj
    (a33, 2.0 |> XlObj.ofFloat, -2.0 |> XlObj.ofFloat, missing, missing)
    |> Slice.xlSlice
    |> Array2D.shouldEqual a33.[1..1, *]

  [<Test>]
  member __.``select center``() =
    let a33 = [[11; 12; 13]; [21;22;23]; [31; 32; 33]] |> toArrayObj
    (a33, 2.0 |> XlObj.ofFloat, -2.0 |> XlObj.ofFloat, 2.0 |> XlObj.ofFloat, -2.0 |> XlObj.ofFloat)
    |> Slice.xlSlice
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
      ] |> toArrayObj
    (a56, 3.0 |> XlObj.ofFloat, 4.0 |> XlObj.ofFloat, 1.0 |> XlObj.ofFloat, -2.0 |> XlObj.ofFloat)
    |> Slice.xlSlice
    |> Array2D.shouldEqual a56.[2..3, 0..4]
