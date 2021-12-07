namespace WldMr.Excel


module MeasureType =
  [<Measure>] type xlCellT
  [<MeasureAnnotatedAbbreviation>] type obj<[<Measure>] 'm> = obj

open MeasureType

type xlObj = obj<xlCellT>


#nowarn "42"

[<AutoOpen>]
module TagOps =
  let inline private cast<'a, 'b> (a : 'a) : 'b = (# "" a : 'b #)

  [<RequireQualifiedAccess>]
  module XlObj =
    let toObj (x : xlObj) : obj = cast x

    [<RequireQualifiedAccess>]
    module Unsafe =
      let ofObj (x : obj) : obj<'m> = cast x

  [<RequireQualifiedAccess>]
  module XlObjRange =
    [<RequireQualifiedAccess>]
    module Unsafe =
      let ofObj2d (x : obj[,]) : xlObj[,] = cast x
