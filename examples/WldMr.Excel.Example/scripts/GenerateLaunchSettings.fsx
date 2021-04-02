
#r "nuget: FSharpPlus"
#r "nuget: System.Text.Json"  // can nuget be avoided without adding the reference to the project?

open System.IO
open FSharpPlus
open System.Text.Json


type Arch = X64 | X86
module Arch =
  let toSuffix = function X64 -> "64" | _ -> ""

// TODO: FIX me
let findExcel () =
  if File.Exists "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" then
    "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" |> Some
  else
    None

// TODO: improve me
let findArch path = if path |> String.startsWith "C:\\Program Files" then X64 else X86

let launchSettingsContent excelPath projectName args =
  sprintf """{
  "profiles": {
    "%s": {
      "commandName": "Executable",
      "executablePath": "%s",
      "commandLineArgs": "/x %s"
    }
  }
}
""" projectName excelPath args 

let getPaths () =
  let projectPath = __SOURCE_DIRECTORY__ |> Path.GetDirectoryName
  let projectName = projectPath |> Path.GetFileName
  let excelPath = findExcel () |> Option.defaultValue "UNKNOWN_PATH_TO_EXCEL"
  let relativeXllPath = "\\bin\\Debug\\net461"
  let xllName = projectName + (excelPath |> findArch  |> Arch.toSuffix) + ".xll"
  {|
    ProjectPath = projectPath
    ProjectName = projectName
    ExcelPath = excelPath
    XllPath = projectPath + relativeXllPath + "\\" + xllName
  |}


let fileContent () = 
  let escapeString (s: string) = JsonEncodedText.Encode(s).ToString()
  let paths = getPaths()
  launchSettingsContent (escapeString paths.ExcelPath) (escapeString paths.ProjectName) (escapeString paths.XllPath)


let outputPath = """\Properties\launchSettings.json"""

let writeLaunchSettings () =
  let paths = getPaths ()
  paths.ProjectPath + "\\Properties" |> System.IO.Directory.CreateDirectory |> ignore
  let templateFullPath = paths.ProjectPath + outputPath
  printfn "Writing to %s" templateFullPath
  File.WriteAllText(templateFullPath, fileContent())
  "It seems the script terminated without error!" + "\n" +
    "The file should have been created at:" + "\n" +
    $"{templateFullPath}" |> printfn "%s"
