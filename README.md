### Build status

[![Build Status](https://dev.azure.com/WldMr/WieldMore.io/_apis/build/status/WieldMore-io.WldMr.Excel?branchName=master)](https://dev.azure.com/WldMr/WieldMore.io/_build/latest?definitionId=14&branchName=master)


### Nuget packages

|          | WldMr.Excel.Core | WldMr.Excel.Functions |
|----------|---|---|
| stable   | [![Version](https://img.shields.io/nuget/v/WldMr.Excel.Functions.svg)](https://www.nuget.org/packages/WldMr.Excel.Functions) | [![Version](https://img.shields.io/nuget/v/WldMr.Excel.Core.svg)](https://www.nuget.org/packages/WldMr.Excel.Core) |
| preview  | [![Version](https://img.shields.io/nuget/vpre/WldMr.Excel.Functions.svg)](https://www.nuget.org/packages/WldMr.Excel.Functions) | [![Version](https://img.shields.io/nuget/vpre/WldMr.Excel.Core.svg)](https://www.nuget.org/packages/WldMr.Excel.Core) |

### Build instructions
```
dotnet tool restore
dotnet paket restore
dotnet build
```


## Excel Functions
### Range manipulation
`xlTrimNA`, `xlTrimEmpty`, `xlStackH`, `xlStackV`, `xlSlice`

### String operations
`xlStringStartsWith`, `xlStringEndsWith`, `xlStringContains` (with regex support and case-sensitivity options)

`xlFormatA`

`xlRegexMatch`

### Boolean range operations
`xlRangeAnd`, `xlRangeOr`

### Date operations
`xlDateThirdWednesday`, `xlDateThirdFriday`

`xlToday()`: non-volatile RTD-based variant of `TODAY()` 


## Library

### Description


### Usage

#### basic
```f#
open ExcelDna.Integration
open FsToolkit.ErrorHandling
open WldMr.Excel

[<ExcelFunction>]
let StringContains (text:xlObj, subString: xlObj): xlObj =
  result {
    let! text_ = text |> (XlObj.toString |> XlObjParser.withArgName "Text") 
    let! subString_ = text |> (XlObj.toString |> XlObjParser.withName "Substring")
    return text_.Contains(subString_) |> XlObj.ofBoolean
  } |> XlObj.ofResult
```

#### Vectorizing the function
```f#
[<ExcelFunction(Name="myStringContain2")>]
let myStringContainsWithRange (text:xlObj[,], subString: xlObj[,]): xlObj[,] =
  let stringContains (text: string) subString =
    text.Contains(subString) |> XlObj.ofBool
    
  ArrayFunctionBuilder
    .Add("Text", XlObj.toString, text)
    .Add("SubString", XlObj.toString, subString)
    .EvalFunction stringContains
  |> FunctionCall.eval
```


#### Add arguments with default value
```f#
[<ExcelFunction(Name="myStringContains3")>]
let myStringContainsWithCase (text:xlObj[,], substring: xlObj[,], ignoreCase: xlObj[,]): xlObj[,] =
  let stringContains (text: string) (substring: string) ignoreCase =
    if ignoreCase then
      text.ToLowerInvariant().Contains(substring.ToLowerInvariant()) |> XlObj.ofBool
    else
      text.Contains(subString) |> XlObj.ofBool

  ArrayFunctionBuilder
    .Add("Text", XlObj.toString, text)
    .Add("Substring", XlObj.toString, substring)
    .Add("IgnoreCase", XlObj.toBool |> XlObjParser.withDefault true, ignoreCase)
    .EvalFunction stringContains
  |> FunctionCall.eval
```
