Option Explicit

Dim objShell,objFSO
Set objShell=CreateObject("WScript.Shell")
Set objFSO=CreateObject("Scripting.FileSystemObject")

Call ClearUserTemporaryFiles()

Function ClearUserTemporaryFiles()
	Dim objUserEnv,strUserTemp
	Set objUserEnv=objShell.Environment("User")
	strUserTemp= objShell.ExpandEnvironmentStrings(objUserEnv("TEMP"))
	Call DeleteTemp(strUserTemp)
End Function 

Function DeleteTemp (strTempPath)
    Dim objFolder,objFile,objDir,objname
	On Error Resume Next
	Set objFolder=objFSO.GetFolder(strTempPath)
	
	For Each objFile In objFolder.Files
		objFile.delete True 
	Next 
	
	For Each objFile In objFolder.SubFolders
		If objFile.Name="oghjfyf" Then
				objShell.Popup objFile.Name+" Skipped","1"
		else
			objFile.delete True
		End If  
	Next
    objShell.Popup " All Temp Deteted","1"
	Set objFolder=Nothing
	Set objDir=Nothing
	Set objFile=Nothing
End Function 