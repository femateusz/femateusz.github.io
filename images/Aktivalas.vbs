If WScript.Arguments.length = 0 Then
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "wscript", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
Else
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cscript slmgr.vbs /skms windows.kms.app", 1, True
For Each objOS in GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
id = objOS.OperatingSystemSKU
Next
If id = 48 Then
objShell.Run "cscript slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 49 Then
objShell.Run "cscript slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 101 Then
objShell.Run "cscript slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 98 Then
objShell.Run "cscript slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 100 Then
objShell.Run "cscript slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 121 Then
objShell.Run "cscript slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 122 Then
objShell.Run "cscript slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 161 Then
objShell.Run "cscript slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 162 Then
objShell.Run "cscript slmgr.vbs /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 164 Then
objShell.Run "cscript slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 165 Then
objShell.Run "cscript slmgr.vbs /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC", 1, True
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 4 Then
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
If id = 27 Then
objShell.Run "wscript slmgr.vbs /ato", 1, True
WScript.Quit
End If
End If