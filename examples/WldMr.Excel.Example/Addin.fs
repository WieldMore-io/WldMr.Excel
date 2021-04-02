module WldMr.Excel.Example.Addin

open ExcelDna.Integration
open ExcelDna.IntelliSense


type AddIn() =
  interface IExcelAddIn with
    member _.AutoOpen() =
      IntelliSenseServer.Install()

    member _.AutoClose() =
      IntelliSenseServer.Uninstall()
  end
