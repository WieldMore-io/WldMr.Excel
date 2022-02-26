namespace WldMr.Excel.Core.Extensions


[<RequireQualifiedAccess>]
module Array2D =
  /// <summary>
  /// flattens the array as an array of rows, each row being an array
  /// simplify further processing (although potentially slower)
  /// </summary>
  let flatten array2d =
    Array.init (Array2D.length1 array2d)
      (fun i ->
        Array.init (Array2D.length2 array2d)
          (fun j -> array2d.[i, j]))


[<RequireQualifiedAccess>]
module Async =
  let map f (a: Async<_>) =
    async {
      let! v = a
      return f v
    }
