Option Explicit

Dim objShell,objFSO,c,arr(10)
c=0
Set objShell=CreateObject("WScript.Shell")
Set objFSO=CreateObject("Scripting.FileSystemObject")

Call ClearSystemTemporaryFiles()
Call ClearUserTemporaryFiles()
Call DeleteTempInternetFiles()
Call ClearIECache()
'Call ClearChromeCache()
Call WriteInText()


Function ClearSystemTemporaryFiles()
	Dim objSysEnv,strSysTemp	
	Set objSysEnv=objShell.Environment("System")
	strSysTemp= objShell.ExpandEnvironmentStrings(objSysEnv("TEMP"))
	arr(c)= strSysTemp
	c=c+1
	Call DeleteTemp(strSysTemp)
End Function 

Function ClearUserTemporaryFiles()
	Dim objUserEnv,strUserTemp
	Set objUserEnv=objShell.Environment("User")
	strUserTemp= objShell.ExpandEnvironmentStrings(objUserEnv("TEMP"))
	arr(c)=strUserTemp
	c=c+1
	Call DeleteTemp(strUserTemp)
End Function 

Function DeleteTemp (strTempPath)
    Dim objFolder,objFile,objDir
	On Error Resume Next
	Set objFolder=objFSO.GetFolder(strTempPath)
	For i=0 To 5 Step 1
		For Each objFile In objFolder.Files
			If Not InStr (objFile.Name,"Bookmarks") Then
				objFile.delete True
			End If 
		Next
    		For Each objDir In objFolder.SubFolders
    			If Not InStr (objDir.Name,"Default") Then
					objDir.delete True
				End If 
        	Next
	Next
	Set objFolder=Nothing
	Set objDir=Nothing
	Set objFile=Nothing
End Function 


Function DeleteTempInternetFiles()
	Dim OSType,TempInternetFiles
	OSType=FindOSType()
	If OSType="Windows 10" Then 
		TempInternetFiles=GetUserProfile() & "\AppData\Local\Microsoft\Windows\INetCache"
	ElseIf OSType="Windows 7" Or OSType="Windows Vista" Then
		TempInternetFiles=GetUserProfile() & "\AppData\Local\Microsoft\Windows\Temporary Internet Files"
	ElseIf  OSType="Windows 2003" Or OSType="Windows XP" Then
		TempInternetFiles=GetUserProfile() & "\Local Settings\Temporary Internet Files"
	End If
	arr(c)=TempInternetFiles
	c=c+1
	DeleteTemp(TempInternetFiles)
	'this is also to delete Content.IE5 in Internet Temp files
	TempInternetFiles=TempInternetFiles & "\Content.IE5"
	arr(c)=TempInternetFiles
	c=c+1
	DeleteTemp(TempInternetFiles)
End Function

Function FindOSType()
    Dim objWMI, objItem, colItems,OSVersion, OSName,ComputerName
   	ComputerName="."
    'Get the WMI object and query results
    Set objWMI = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
    'Get the OS version number (first two) and OS product type (server or desktop) 
    For Each objItem in colItems
        OSVersion = Left(objItem.Version,3)
		'msgbox (OSVersion )
    Next
    Select Case OSVersion
		Case "10."
            OSName = "Windows 10"
        Case "6.1"
            OSName = "Windows 7"
        Case "6.0" 
            OSName = "Windows Vista"
        Case "5.2" 
            OSName = "Windows 2003"
        Case "5.1" 
            OSName = "Windows XP"
        Case "5.0" 
            OSName = "Windows 2000"
   	End Select
    FindOSType = OSName
    Set colItems = Nothing
    Set objWMI = Nothing
End Function 

Function GetUserProfile()
GetUserProfile = objShell.ExpandEnvironmentStrings("%userprofile%")
End Function 

Function ClearIECache()
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8",2
	'To clear browsing cookies
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2",2
	'To Clear Browsing History
	objShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1",2
End Function 


Function ClearChromeCache()
	Dim Path,File
	Path = GetUserProfile()& "\AppData\Local\Google\Chrome\User Data\"
	arr(c)=Path
	c=c+1	
    	DeleteTemp(Path)
End Function 

Function WriteInText()
	Dim i,objFileToWrite
	Set objFileToWrite =objFSO.OpenTextFile (GetUserProfile()&"\MyFolder\DeletedPaths.txt",2,True)
	For i=0 To UBound(arr) Step 1 
		objFileToWrite.WriteLine(arr(i))
	Next 
	objShell.Popup "All Cache Cleared",1,"Alert"
End Function

Set objFSO=Nothing
WScript.Quit