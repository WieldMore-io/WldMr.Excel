namespace WldMr.Excel

open ExcelDna.Integration
open System
open System.Threading
open WldMr.Excel.Core.Extensions

#nowarn "42"

module AsyncFunctionCall =
  let retrievingString = "#Retrieving"

  // from ExcelDNA's Distribution/Samples/Async/FsAsync.dna
  [<Obsolete>]
  let excelObserve functionName parameters observable =
    let obsSource =
      ExcelObservableSource(
        fun () ->
        { new IExcelObservable with
          member _.Subscribe observer =
            // Subscribe to the F# observable
            Observable.subscribe (fun value -> observer.OnNext (value)) observable
        })
    ExcelAsyncUtil.Observe (functionName, parameters, obsSource)


  let private excelObservableSource async =
    ExcelObservableSource(fun () ->
      { new IExcelObservable with
          member _.Subscribe observer =
            // make something like CancellationDisposable
            let cts = new CancellationTokenSource ()
            let disposable = { new IDisposable with member _.Dispose () = cts.Cancel () }
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


  module Cell =
    let wrapAsync functionName parameters (async: Async<xlObj>): xlObj =
      try
        match ExcelAsyncUtil.Observe (functionName, parameters, excelObservableSource async) with
        | oneObj ->
            match (%oneObj: xlObj) with
            | ExcelNA _ -> retrievingString |> XlObj.ofString
            | o -> o
      with
        | e -> $"{e.Message} ({e.GetType()})" |> XlObj.ofErrorMessage


  module Range =
    let wrapAsync functionName parameters (async: Async<xlObj[,]>): xlObj[,] =
      try
        match ExcelAsyncUtil.Observe (functionName, parameters, excelObservableSource async) with
        | :? (obj[,]) as a -> (# "" a : xlObj[,] #)
        | oneObj ->
            match (%oneObj: xlObj) with
            | ExcelNA _ -> retrievingString |> XlObj.ofString |> XlObjRange.ofCell
            | o -> o |> XlObjRange.ofCell
      with
        | e -> $"{e.Message} ({e.GetType()})" |> XlObj.ofErrorMessage |> XlObjRange.ofCell
