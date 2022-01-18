module WldMr.Excel.Example.Tutorial

open ExcelDna.Integration
open FsToolkit.ErrorHandling
open WldMr.Excel

[<ExcelFunction>]
let myStringContains1 (text:xlObj, substring: xlObj): xlObj =
  result {
    let! text_ = text |> (XlObj.toString |> XlObjParser.withArgName "Text")
    let! substring_ = substring |> (XlObj.toString |> XlObjParser.withArgName "Substring")
    return
      text_.Contains(substring_)
      |> XlObj.ofBool
  } |> XlObj.ofResult


[<ExcelFunction(Name="myStringContain2")>]
let myStringContainsWithRange (text:xlObj[,], subString: xlObj[,]): xlObj[,] =
  let stringContains (text: string) (subString: string) =
    text.Contains(subString) |> XlObj.ofBool

  ArrayFunctionBuilder
    .Add("Text", XlObj.toString, text)
    .Add("SubString", XlObj.toString, subString)
    .EvalFunction stringContains
  |> FunctionCall.eval


[<ExcelFunction(Name="myStringContains3")>]
let myStringContainsWithCase (text:xlObj[,], subString: xlObj[,], ignoreCase: xlObj[,]): xlObj[,] =
  let stringContains (text: string) (subString: string) ignoreCase =
    if ignoreCase then
      text.ToLowerInvariant().Contains(subString.ToLowerInvariant()) |> XlObj.ofBool
    else
      text.Contains(subString) |> XlObj.ofBool

  ArrayFunctionBuilder
    .Add("Text", XlObj.toString, text)
    .Add("SubString", XlObj.toString, subString)
    .Add("IgnoreCase", XlObj.toBool |> XlObjParser.withDefault true, ignoreCase)
    .EvalFunction stringContains
  |> FunctionCall.eval
