namespace WldMr.Excel

open ExcelDna.Integration
open System
open System.Threading
open WldMr.Excel.Core.Extensions

#nowarn "42"


module AsyncFunctionCall =
  let retrievingString = "#Retrieving"


  let private excelObservableFromObservable (event: IObservable<_>) =
    ExcelObservableSource(fun () ->
      { new IExcelObservable with
          member _.Subscribe observer =
            Observable.subscribe observer.OnNext event
      })


  let private excelObservableFromAsync async =
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


  module internal CellInternal =
    let private wrapCommon functionName parameters excelObservable: xlObj =
      try
        match ExcelAsyncUtil.Observe (functionName, parameters, excelObservable) with
        | oneObj ->
            match (%oneObj: xlObj) with
            | ExcelNA _ -> retrievingString |> XlObj.ofString
            | o -> o
      with
        | e -> $"{e.Message} ({e.GetType()})" |> XlObj.ofErrorMessage

    let wrapAsync functionName parameters (async: Async<xlObj>): xlObj =
      excelObservableFromAsync async
      |> wrapCommon functionName parameters


    let wrapObservable functionName parameters (observable: IObservable<xlObj>): xlObj =
      excelObservableFromObservable observable
      |> wrapCommon functionName parameters


  module internal RangeInternal =
    let private wrapCommon functionName parameters excelObservable: xlObj[,] =
      try
        match ExcelAsyncUtil.Observe (functionName, parameters, excelObservable) with
        | :? (obj[,]) as a -> (# "" a : xlObj[,] #)
        | oneObj ->
            match (%oneObj: xlObj) with
            | ExcelNA _ -> retrievingString |> XlObj.ofString |> XlObjRange.ofCell
            | o -> o |> XlObjRange.ofCell
      with
        | e -> $"{e.Message} ({e.GetType()})" |> XlObj.ofErrorMessage |> XlObjRange.ofCell

    let wrapAsync functionName parameters (async: Async<xlObj[,]>): xlObj[,] =
      excelObservableFromAsync async
      |> wrapCommon functionName parameters


    let wrapObservable functionName parameters (observable: IObservable<xlObj[,]>): xlObj[,] =
      excelObservableFromObservable observable
      |> wrapCommon functionName parameters


  module Cell =
    let fromAsync name hash a =
      fun () ->
        a
        |> Async.map XlObj.ofResult
        |> CellInternal.wrapAsync name hash

    let fromObservable name hash a =
      fun () -> CellInternal.wrapObservable name hash a


  module Range =
    let fromAsync name hash a =
      fun () ->
        a
        |> Async.map XlObjRange.ofResult
        |> RangeInternal.wrapAsync name hash

    let fromObservable name hash a =
      fun () -> RangeInternal.wrapObservable name hash a
