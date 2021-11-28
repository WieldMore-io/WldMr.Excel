namespace WldMr.Excel.Utilities

open ExcelDna.Integration
open System
open System.Threading

#nowarn "42"

module ExcelAsync =
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

  module Cell =
    let wrapAsync functionName parameters (async: Async<objCell>): objCell =
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
      ExcelAsyncUtil.Observe (functionName, parameters, obsSource) |> (~%)

    let wrapAsyncResult functionName parameters asyncResult =
      let asyncFun =
        async {
          let! result = asyncResult
          return result |> XlObj.ofResult
        }
      wrapAsync functionName parameters asyncFun


  module Range =
    let wrapAsync functionName parameters (async: Async<objCell[,]>): objCell[,] =
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
      let eaRes = ExcelAsyncUtil.Observe (functionName, parameters, obsSource)
      match eaRes with
      | :? (obj[,]) as a -> (# "" a : objCell[,] #)
      | oneObj -> (%oneObj: objCell) |> Array2D.create 1 1

    let wrapAsyncResult functionName parameters asyncResult =
      let asyncFun =
        async {
          let! result = asyncResult
          return result |> XlObjRange.ofResult
        }
      wrapAsync functionName parameters asyncFun
