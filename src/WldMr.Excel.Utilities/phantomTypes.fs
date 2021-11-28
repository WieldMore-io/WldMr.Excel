namespace WldMr.Excel.Utilities



[<Measure>] type xlCellT
//[<Measure>] type xlRangeT

[<MeasureAnnotatedAbbreviation>] type obj<[<Measure>] 'm> = obj

#nowarn "42"
module private Unsafe =
  let inline cast<'a, 'b> (a : 'a) : 'b =
    (# "" a : 'b #)

type XlTypes =
  static member inline tag<[<Measure>]'m> (x : obj) : obj<'m> = Unsafe.cast x
  static member inline untag<[<Measure>]'m> (x : obj<'m>) : obj = Unsafe.cast x
  static member inline cast<[<Measure>]'m1, [<Measure>]'m2> (x : obj<'m1>) : obj<'m2> = Unsafe.cast x


//  static member inline castToRange (x : obj<xlCellT>) : obj<xlRangeT> = Unsafe.cast x

type objCell = obj<xlCellT>
//type objRange = obj<xlRangeT>

[<AutoOpen>]
module Operators =

  let inline private _cast< ^TC, ^xm, ^xn when (^TC or ^xm or ^xn) : (static member cast : ^xm -> ^xn)> (x : ^xm) =
      ((^TC or ^xm or ^xn) : (static member cast : ^xm -> ^xn) x)

  // NB the particular infix operator shadows the rarely used quotation splicing operator
  /// Infix operator used for tagging, untagging, or casting units of measure
  let inline (~%) (x : obj) : obj<_> = _cast<XlTypes, obj, obj<_>> x

