// Running this file will generate suitable (?) default 
// for starting a debug session with Visual Studio
// To run this script in Visual Studio:
//   Select all the text in this file (ie Ctrl-A)
//   Send it to F# interactive (Alt-Enter)
//   Watch the output in the F# Interactive window

#load "scripts/GenerateLaunchSettings.fsx"
open GenerateLaunchSettings

writeLaunchSettings()
