namespace WldMr.Excel.Functions

[<AutoOpen>]
module FSharpCoreExtensions =
  [<RequireQualifiedAccess>]
  module String =
    let inline startsWith (prefix: string) (s: string) = s.StartsWith prefix
    let inline toUpper (s: string) = s.ToUpper()
    let inline toLower (s: string) = s.ToLower()
    let inline contains (sub: string) (s: string) = s.Contains sub
    let inline isEmpty (s: string) = s.Length = 0

  [<RequireQualifiedAccess>]
  module Result =
    let protect f x =
      try
        f x |> Ok
      with e -> e |> Error


