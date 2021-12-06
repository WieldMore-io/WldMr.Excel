namespace WldMr.Excel.Core.Extensions


[<RequireQualifiedAccess>]
module Array2D =
  /// <summary>
  /// flattens the array as a list of rows, each row being a list
  /// simplify further processing (although potentially slower)
  /// </summary>
  let flatten array2d =
    [ for x in [ 0 .. (Array2D.length1 array2d) - 1 ] do
      [ for y in [ 0 .. (Array2D.length2 array2d) - 1 ] do
        yield array2d.[x, y] ] ]

[<RequireQualifiedAccess>]
module Async =
  let map f (a: Async<_>) =
    async {
      let! v = a
      return f v
    }
