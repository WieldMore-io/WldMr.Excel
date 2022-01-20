module WldMr.Excel.Example.Addin

open ExcelDna.Integration
#if NETFRAMEWORK
open ExcelDna.IntelliSense
#endif

type AddIn() =
  interface IExcelAddIn with
    member _.AutoOpen() =
      #if NETFRAMEWORK
      IntelliSenseServer.Install()
      #endif
      ()

    member _.AutoClose() =
      #if NETFRAMEWORK
      IntelliSenseServer.Uninstall()
      #endif
      ()
  end
