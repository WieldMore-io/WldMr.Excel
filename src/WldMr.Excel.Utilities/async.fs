namespace WldMr.Excel.Utilities

open ExcelDna.Integration
open System
open System.Threading


module ExcelAsync =
  [<Obsolete("Use wrapAsync or excelObserve")>]
  let wrap<'T> (funName: string) (hash: int) (errMsg: obj) (waitMsg: obj) (f: unit -> obj): obj =
    let asyncResult = ExcelAsyncUtil.Run(funName, hash, new ExcelFunc(f))
    match asyncResult with
    | null ->
        errMsg
    | ExcelError(ExcelError.ExcelErrorNA) ->
        waitMsg
    | :? System.Tuple<string, string> as tu ->
        let t1, t2 = tu
        (if t1 = null then t2 else t1) |> box
    | _ ->
        asyncResult


  //
  // from ExcelDNA's Distribution/Samples/Async/FsAsync.dna
  //

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

  let wrapAsync functionName parameters async =
    let obsSource =
      ExcelObservableSource(
        fun () ->
        { new IExcelObservable with
            member __.Subscribe observer =
              // make something like CancellationDisposable
              let cts = new CancellationTokenSource ()
              let disposable = { new IDisposable with member __.Dispose () = cts.Cancel () }
              // Start the async computation on this thread
              Async.StartWithContinuations
                ( computation= async,
                  continuation= ( fun result -> observer.OnNext(result); observer.OnCompleted () ),
                  exceptionContinuation= observer.OnError ,
                  cancellationContinuation= ( fun _ -> observer.OnCompleted () ),
                  cancellationToken= cts.Token
                )
              // return the disposable
              disposable
        })
    ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
