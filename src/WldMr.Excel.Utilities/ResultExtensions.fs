namespace WldMr

module Result =
  let value  = function
    | Ok v -> v
    | Error e ->
      match box e with
      | :? string as s -> failwith s
      | :? exn as e -> raise e
      | e -> failwith (e.ToString())

  let toOption = function
    | Ok v -> Some v
    | _ -> None

  let protect f x =
    try
      f x |> Ok
    with
    | e -> e |> Error

  let inline isOk r =
    match r with
    | Ok _ -> true
    | _ -> false

