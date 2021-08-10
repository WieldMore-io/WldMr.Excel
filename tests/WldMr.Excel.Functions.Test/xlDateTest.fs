namespace xlDate

open NUnit.Framework
open FsUnit

open WldMr.Excel.Date.NthWeekdayOfMonth

[<TestFixture>]
type ``nth Weekday of the month``() =
  [<Test>]
  member __.``is correct for n = 1..4``() =
    let years = [2000 .. 2013]
    let months = [1 .. 12]

    for y in years do
      for m in months do
        for wd in [0 .. 6] do
          for wn in [1 .. 4] do
            let d = nThDayOfWeekForMonth wn wd y m
            d.DayOfWeek |> should equal (enum wd: System.DayOfWeek)
            d.Day |> should be (lessThanOrEqualTo (7 * wn))
            d.Day |> should be (greaterThanOrEqualTo (7 * (wn - 1)))
          
          
