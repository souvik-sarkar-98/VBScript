Set objShell=CreateObject("WScript.Shell")

Call ClearIECache()

Function ClearIECache()
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8",2
	'To clear browsing cookies
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2",2
	'To Clear Browsing History
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1",2
End Function