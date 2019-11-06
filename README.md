# VBA Analytics

Way of tracking code analytics and metrics. Useful when multiple users are running code within a shared network.

![graph animation](./assets/hero2.gif)

---

## Other Helpful Resources

- [www.roberttodar.com](https://www.roberttodar.com/) About me and my background and some of my other projects.
- [Style Guide](https://github.com/todar/VBA-Style-Guide) A guide for writing clean VBA code. Notes on how to take notes =)
- [Boilerplate](https://github.com/todar/VBA-Boilerplate) A Boilerplate that contains a bunch of helper libraries such as JSON tools, Code Analytics, LocalStorage, Unit Testing, version control and local network distribution, userform events, and more!
- [Arrays](https://github.com/todar/VBA-Arrays) An Array function library, functions that mimic JavaScript Arrays.
- [Strings](https://github.com/todar/VBA-Strings) String function library. `ToString`, `Inject`, `StringSimilarity`, and more.
- [Userform EventListener](https://github.com/todar/VBA-Userform-EventListener) Listen to events such as `mouseover`, `mouseout`, `focus`, `blur`, and more.

---

## Required

Import `AnalyticsTracker.cls` and `JSON.cls` into your project.

Set a reference to the `Microsoft Scripting Runtime` as both of these classes use `Scripting.Dictionary` and `Scripting.FileSystemObject`.

---

## Usage

Initialize a new instance of `AnalyticsTracker` and call the method `TrackStats` to start tracking. Place this near the top of your code or from the point you want to start tracking.

To complete the tracking you can call `FinalizeStats`. Note that this is optional as you can let the `Class_Terminate` event fire that method itself.

Optionaly you can add stats to the main thread using `AddStat` or even add your own custom events using `LogEvent`.

```vb
'/**
' * Sample of how to track and use Analytics class.
' *
' * @ref {Class Module} AnalyticsTracker
' */
Private Sub howToTrackAnalytics()
    ' This tracks to a JSON file and the immediate window.
    ' To be effecent this appends to the text file.
    ' Because of this the JSON file is missing the outer array
    ' brackets []. Also includes a comma after each object {},
    ' So to use this as JSON you must edit those two things.
    Dim analytics As New AnalyticsTracker

    ' You can track standard stats for code use!
    ' This collects codeName, username, date, time, timesaved, runtime
    analytics.TrackStats "test", 5

    ' You can also optionally add custom stats to the main thread.
    analytics.AddStat "customStat", "I'm custom!"

    ' You also have the ability to log your own custom events. This by default
    ' still adds standard stats date, time, username.
    analytics.LogEvent "onCustom", "name", "Robert", "age", 31

    ' Optional! You can either call this function or let the
    ' terminate event in the class to run it. Really the only
    ' function that needs to be called in order to collect stats
    ' is `TrackStats` or `AddStat`. Put that at the top of a function
    ' and it will log all the basic stats.
    analytics.FinalizeStats
End Sub
```

Here is an example showing the minimum amount of code needed to track basic stats.

```vb
Private Sub minimumNeededToTrackAnalytics()
    Dim analytics As New AnalyticsTracker
    analytics.TrackStats "test"
    ' Do Something...
End Sub
```

---

## The data

Immediate Window dislays a formatted version of the JSON.

```
Tracking Analitics For test
{
  "event": "onUse",
  "id": "E7846A245BC04750813A47B3020EB84A",
  "codeName": "test",
  "username": "rtoda02",
  "date": "11/6/2019",
  "time": "1:37:53 PM",
  "runtimeSeconds": 0
}
```

Analytics.json file is appended. This means the outer brackets `[]` are missing from the JSON file and the last added record has a comma that needs to be removed.

```json
{
  "event": "onUse",
  "id": "E7846A245BC04750813A47B3020EB84A",
  "codeName": "test",
  "username": "rtoda02",
  "date": "11/6/2019",
  "time": "1:37:53 PM",
  "runtimeSeconds": 0
},
```

To use this JSON add the array brackets `[]` and remove the last record comma.

```json
[
  {
    "event": "onUse",
    "id": "E7846A245BC04750813A47B3020EB84A",
    "codeName": "test",
    "username": "rtoda02",
    "date": "11/6/2019",
    "time": "1:37:53 PM",
    "runtimeSeconds": 0
  }
]
```
