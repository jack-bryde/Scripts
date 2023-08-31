' This script changes the default browser program set by group policy.
' Uses send keys, impersonating a human PC user.
'

' As scheduled to commence on log-in, set initial wait time to reduce risk of
' interruption from Outlook/Teams etc.
WScript.Sleep 60000 '1 minute

Set WshShell = WScript.CreateObject("WScript.Shell")
' Open default settings window
WshShell.Run "ms-settings:defaultapps"
' Wait until open
WScript.Sleep 5000

' Navigate down to the browser setting
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"

' Open the browser choice menu
WshShell.SendKeys " "
' Wait until open
WScript.Sleep 500
' Navigate to Chrome and select
WshShell.SendKeys "{TAB}"
WshShell.SendKeys " "

' Close the default program window
WScript.Sleep 2000 ' Wait 2 seconds
WshShell.SendKeys "%{F4}"

WScript.Quit
