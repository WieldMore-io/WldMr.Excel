namespace WldMr

module Option =
  let toResult errorFun = function
    | Some v -> Ok v
    | None -> Error (errorFun ())

  let inline ofTry (b, t) =
    if b then Some t else None
