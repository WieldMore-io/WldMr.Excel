open System.IO


let excelKey = """SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"""
let excelPath = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(excelKey).GetValue("").ToString()


type Arch = X64 | X86

module Arch =
  let toSuffix = function X64 -> "64" | X86 -> ""

  let GetBinaryType (path:string) =
    let at pos (reader:BinaryReader) = 
      reader.BaseStream.Seek(pos, SeekOrigin.Begin) |> ignore |> reader.ReadInt32
    use reader = new BinaryReader(File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
    let peOffset = reader |> at 60L
    let optHeaderMagic = reader |> at (int64 (peOffset + 24)) |> int16
    match optHeaderMagic with | 0x20Bs -> X64 | 0x10Bs -> X86 | _ -> failwith "Unknown arch"


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
  let relativeXllPath = "\\bin\\Debug\\net461"
  let xllName = projectName + (excelPath |> Arch.GetBinaryType |> Arch.toSuffix) + ".xll"
  {|
    ProjectPath = projectPath
    ProjectName = projectName
    ExcelPath = excelPath
    XllPath = projectPath + relativeXllPath + "\\" + xllName
  |}


let fileContent () = 
  let escapeString (s: string) = System.Web.HttpUtility.JavaScriptStringEncode(s)
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
