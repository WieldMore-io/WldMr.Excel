namespace WldMr

module List =
  let headResult = function
    | [] -> Result.Error "Empty list"
    | x::l -> Result.Ok x
