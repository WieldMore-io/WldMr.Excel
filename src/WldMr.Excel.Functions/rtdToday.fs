namespace WldMr.Excel.Rtd.Today

open ExcelDna.Integration
open System
open System.Collections.Generic
open System.Runtime.InteropServices
open System.Threading
open ExcelDna.Integration.Rtd

module RtdTodayServer =
  [<Literal>]
  let progId = "WldMr.Today"

[<ComVisible(true)>]  // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
[<ProgId(RtdTodayServer.progId)>]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
type RtdTodayServer() =
  inherit ExcelRtdServer()

  // Using a System.Threading.Time which invokes the callback on a ThreadPool thread
  // (normally that would be dangerous for an RTD server, but ExcelRtdServer is thread-safe)
  let mutable timer = null
  let topics = List<ExcelRtdServer.Topic>()

  override this.ServerStart() =
    let now = DateTime.Now
    let n = now.AddDays 1.0
    let r = DateTime(n.Year, n.Month, n.Day)
    let d = (int) (r - now).TotalMilliseconds
    timer <- new Timer(this.timer_tick, null, d, 86400 * 1000)
    topics.Clear()
    true

  override _.ServerTerminate() =
    timer.Dispose()

  override _.ConnectData(topic, topicInfo:IList<string>, newValues:byref<bool>): obj =
    topics.Add topic
    DateTime.Today :> obj

  override _.DisconnectData(topic): unit =
    topics.Remove topic |> ignore

  member _.timer_tick(_unused_state_: obj): unit =
    let today = DateTime.Today
    let tillTomorrow = (today.AddDays(1.0) - DateTime.Now).TotalSeconds
    let r = if (tillTomorrow < 5.0) then today.AddDays 1.0 else today
    for topic in topics do
      topic.UpdateValue r

module RtdTodayFunction =
  [<ExcelFunction(Category = "WldMr Date", Description = "Non-volatile version of Today()")>]
  let xlToday(): obj =
    // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
    // Note that the topic information needs at least one string - it's not used in this sample
    XlCall.RTD(RtdTodayServer.progId, null, "")
