namespace WldMr.Excel

open ExcelDna.Integration
open FsToolkit.ErrorHandling

open WldMr.Excel

module FunctionCall =
  let eval f = f ()


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
