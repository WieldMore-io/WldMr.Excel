namespace WldMr.Excel

open ExcelDna.Integration
open FsToolkit.ErrorHandling

open WldMr.Excel


module FunctionCall =
  let retrievingString = "#Retrieving"
  let eval f = f ()

  let runAsExcelAsyncRange name hash a =
    fun () ->
      a
      |> Async.map XlObjRange.ofResult
      |> AsyncFunctionCall.Range.wrapAsync name hash

  let runAsExcelAsyncCell name hash a =
    fun () ->
      a
      |> Async.map XlObj.ofResult
      |> AsyncFunctionCall.Cell.wrapAsync name hash


type FunctionCall() =
  static member catchExceptions(f: unit -> xlObj[,]) =
    fun () ->
      try f () with
      | e -> $"Exception caught: {e.ToString()}." |> XlObj.ofErrorMessage |> XlObjRange.ofCell

  static member catchExceptions(f: unit -> xlObj) =
    fun () ->
      try f () with
      | e -> $"Exception caught: {e.ToString()}." |> XlObj.ofErrorMessage

  static member disableInWizard(f: unit -> xlObj[,]) =
    fun () ->
      if ExcelDnaUtil.IsInFunctionWizard() then
        "disabled in Function Wizard" |> XlObj.ofErrorMessage |> XlObjRange.ofCell
      else
        f ()

  static member disableInWizard(f: unit -> xlObj) =
    fun () ->
      if ExcelDnaUtil.IsInFunctionWizard() then
        "disabled in Function Wizard" |> XlObj.ofErrorMessage
      else
        f ()

  static member fromAsync (f: Async<Result<xlObj, string>>) =
    (fun () -> f |> Async.RunSynchronously) >> XlObj.ofResult

  static member fromAsync (f: Async<Result<xlObj[,], string>>) =
    (fun () -> f |> Async.RunSynchronously) >> XlObjRange.ofResult
