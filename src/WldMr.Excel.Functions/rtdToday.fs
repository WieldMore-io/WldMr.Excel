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


[<ComVisible(true)>]
[<ProgId(RtdTodayServer.progId)>]
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
    let msTillTomorrow = (int) (r - now).TotalMilliseconds |> max 5000
    let waitMs = min (60 * 1000) msTillTomorrow
    timer <- new Timer(this.timer_tick, null, waitMs, 60 * 1000)
    topics.Clear()
    true

  override _.ServerTerminate() =
    timer.Dispose()

  override _.ConnectData(topic, _topicInfo: IList<string>, newValues: byref<bool>): obj =
    topics.Add topic
    DateTime.Today :> obj

  override _.DisconnectData(topic): unit =
    topics.Remove topic |> ignore

  member _.timer_tick(_unusedState: obj): unit =
    let today = DateTime.Today
    let tillTomorrow = (today.AddDays(1.0) - DateTime.Now).TotalSeconds
    let r = if (tillTomorrow < 5.0) then today.AddDays 1.0 else today
    for topic in topics do
      topic.UpdateValue r

module RtdTodayFunction =
  [<ExcelFunction(Category = "WldMr Date", Description = "Non-volatile version of Today()")>]
  let xlToday(): obj =
    // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
    // Note that the topic information needs at least one string - it's not used here
    XlCall.RTD(RtdTodayServer.progId, null, "")
