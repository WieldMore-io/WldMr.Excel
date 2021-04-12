namespace WldMr.Excel.Utilities

open ExcelDna.Integration


module ExcelAsync =
  let wrap<'T> (funName: string) (hash: int) (errMsg: obj) (waitMsg: obj) (f: unit -> obj): obj =
    let asyncResult = ExcelAsyncUtil.Run(funName, hash, new ExcelFunc(f))
    match asyncResult with
    | null -> errMsg
    | ExcelError(ExcelError.ExcelErrorNA) -> waitMsg
    | :? System.Tuple<string, string> as tu ->
        let t1, t2 = tu
        (if t1 = null then t2 else t1) |> box
    | _ -> asyncResult

  let excelObserve functionName parameters observable =
    let obsSource =
        ExcelObservableSource(
            fun () ->
            { new IExcelObservable with
                member __.Subscribe observer =
                    // Subscribe to the F# observable
                    Observable.subscribe (fun value -> observer.OnNext (value)) observable
            })
    ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
