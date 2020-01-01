		Option Explicit
		On error resume next 
		Dim FObj,WshShell,WordObj,ImageFolderPath,LocalDocumentPath,BaseSheredPath,StaffCode

StaffCode="XY59036"	
ImageFolderPath="C:\Users\"&StaffCode&"\Documents\Screenshot\PrtScTemp" 'Local temporary path Screenshort folder
LocalDocumentPath="C:\Users\"&StaffCode&"\Documents\Screenshot\Document" 'Local document path  
BaseSheredPath="\\omsmds001\isd\TCS Testing Team\Souvik Sarkar"  'Enter base Shered path address'

		Set FObj= CreateObject("Scripting.FileSystemObject")
		Set WshShell = CreateObject("WScript.Shell")
		Set WordObj = CreateObject("Word.Application")
		
		Const END_OF_STORY = 6
		Const MOVE_SELECTION = 0
		
		
Call Start()

Function Start()
		Select Case InputBox("DO YOU WANT TO SAVE YOUR LAST SCREENSHOTS?" &vbCrLf&""&vbCrLf&"1--> YES"&vbCrLf& "0-->> NO") 
		Case "1"
			Call NewFile()
		Case "0"
			Call ClearImageFolder()
		Case Else
			WshShell.Popup "WRONG INPUT - PLEASE TRY AGAIN", 2, "ALERT" 
			Call Start()
 	End Select 
End Function 

'Main Action function 	
Function Action()
	Select Case InputBox("1-->> IMPORT SCREENSHOTS TO NEW FILE" &vbCrLf&""&vbCrLf&"22-->ADD SCREENSHOT TO EXISTING FILE"&vbCrLf&""&vbCrLf& "333-->> DELETE SCREENSHOTS") 
		Case "1"
			Call NewFile()
		Case "22"
			Call SelectExistingFile()
		Case "333"
			Call ClearImageFolder()
		Case Else
			WshShell.Popup "WRONG INPUT - PLEASE TRY AGAIN", 2, "ALERT" 
			Call Action()
 	End Select 
End Function


'Main Function to inttiate pasting in new word document
Function NewFile()
		If  OpenMSWordPaste(CreateMSWord(CreateFolder(MonthName(Month(now)) &" "& Day(now),CreateFolder(MonthName(Month(now)) &" "& Year(now),BaseSheredPath)))) Then
				Call ClearImageFolder()
		End If  
End Function

'Function to Create a new folder
Function CreateFolder(ByVal FolderName,ByVal FolderPath)
	If Not FObj.FolderExists(FolderPath &"\"& FolderName) Then
		FObj.CreateFolder(FolderPath &"\"& FolderName)
	End If
	CreateFolder=FolderPath &"\"& FolderName
End Function 

'Function to Create a new Word Document
Function CreateMSWord(ByVal FolderPath)
	Dim DocName,DocObj,var
	var=""
	If FObj.FolderExists(FolderPath) Then 
		DocName=InputBox("ENTER WORD FILE NAME TO 'IMPORT SCREENSHOTS TO NEW FILE'"&vbCrLf&""&vbCrLf&"OR CLICK 'CANCEL' TO ADD SCREENSHOT IN EXISTING FILE ")
		If Not DocName="" Then 
			If InStr (DocName,"/")=0 And InStr (DocName,"\")=0 And InStr (DocName,":")=0 And InStr (DocName,"*")=0 And InStr (DocName,"?")=0 And InStr (DocName,"<")=0 And InStr (DocName,">")=0 And InStr (DocName,"|")=0 And Not DocName=""	Then 		
				If Not FObj.FileExists(FolderPath&"\"&DocName&".docx") Then 
					If CopyContent(DocName) Then 
						Set DocObj = WordObj.Documents.Add()
						DocObj.SaveAs(FolderPath &"\" & DocName&".docx")
						var=FolderPath &"\" & DocName&".docx"
					Else
						WshShell.Popup "EXCEPTION OCCURED : UNABLE TO COPY CONTENT : PLEASE TRY ONCE AGAIN", 5, "ALERT"
						Call NewFile()
					End If 
				Else
					WshShell.Popup "'" &DocName&"'"& " ALREADY EXITS", 2, "ALERT"
					Call NewFile()
				End If
			Else 
				WshShell.Popup "A FILE NAME CAN NOT CONTAIN ANY OF THE FOLLOWING CHARACTERS: "&vbCrLf&"\ / : * ? < > |", 2, "ALERT"
				Call CreateMSWord(FolderPath)
			End If
		Else	
			WshShell.Popup "YOU HAVE CANCELLED 'IMPORT SCREENSHOTS TO NEW FILE' OPTION "&vbCrLf& "PLEASE CHOOSE ACTION ", 2, "ALERT"
			Call Action()
		End If  
	Else  
		WshShell.Popup "EXCEPTION OCCURED : UNABLE TO FIND FOLDER PATH : PLEASE TRY ONCE AGAIN", 2, "ALERT"
		Call NewFile()
	End if  
	CreateMSWord=var					
End Function

'Function to open an existing MS Word documet and paste screenshot in it
Function OpenMSWordPaste(ByVal DocPath)
	Dim objSelection,DocObj,colFiles,objFolder,val,i,img,NumFile,var,c,c1
	c=0 
	c1=0
	If Not IsVoid(ImageFolderPath)  Then
		If FObj.FileExists(DocPath) And Not DocPath="" Then 
			Set DocObj = WordObj.Documents.Open(DocPath)
			WordObj.Visible= True
			Set objSelection = WordObj.Selection
			Set objFolder=FObj.getFolder(ImageFolderPath) 
			WshShell.SendKeys "%{TAB}"  
			NumFile=GetRecentFileName(ImageFolderPath)
			Set colFiles = objFolder.Files
			For i=1 To NumFile step 1
				For Each img in colFiles
					If img.Name =i&".png" Or img.Name =i&".jpg" Or img.Name = i&".jpeg" Or img.Name = i&".bmp"  Then
						If Mid(img, InStrRev (img, ".")+1)="png" Or Mid(img, InStrRev (img, ".")+1)="jpg" Or Mid(img, InStrRev (img, ".")+1)="jpeg" Or Mid(img, InStrRev (img, ".")+1)="bmp"Then 
							objSelection.EndKey END_OF_STORY,MOVE_SELECTION
							objSelection.InlineShapes.AddPicture(img.Path)
							'DocObj = WordObj.ActiveDocument.Save
							c=c+1
						End If 
					End If
				Next
			Next
		Else 
			WshShell.Popup " UNABLE TO OPEN FILE : PLEASE CREATE A NEW FILE ",5,"ALERT"	
			Call NewFile()
		End If 			
		DocObj = WordObj.ActiveDocument.Save
		WordObj.Quit 
		For Each img in colFiles
			If FObj.GetExtensionName(img) ="png" Or FObj.GetExtensionName(img) ="jpg" Or FObj.GetExtensionName(img) ="jpeg"	Or FObj.GetExtensionName(img) ="bmp" Then 
				c1=c1+1
			End If 	
		Next
		If c1=c Then 
			var= True
		Else
			var= False
		End If 
		OpenMSWordPaste=var
	End if 	
End Function 

'Function to copy Temp SS to local storage
Function CopyContent(ByVal DocName)
Dim var
var= False 
	FObj.CopyFolder ImageFolderPath,CreateFolder(MonthName(Month(now)) &" "& Day(now),LocalDocumentPath)&"\"&DocName&"  " &Replace(Time(),":","-")
	var= True 
	CopyContent=var
End Function

'Function to clear temporary Screenshots		
Function ClearImageFolder()
	Dim folder,f
	Set folder=FObj.getFolder(ImageFolderPath)
	For Each f In folder.Files
		if FObj.GetExtensionName(f)="jpg" or FObj.GetExtensionName(f)="png" or FObj.GetExtensionName(f)="jpeg" or FObj.GetExtensionName(f)="bmp" then 
			f.Delete True
		end if
	Next
	WshShell.Popup "SYSTEM IS READY ",3, "ALERT"
End Function

'Main Function to initiate and perform screenshot pasting in existing word document  
Function SelectExistingFile()
	Dim oExec,FileSelected,DocName,DocumentObj 
	WshShell.Popup "PLEASE MAKE SURE THE FILE THAT YOU ARE GOING TO CHOOSE IS ALREADY NOT OPEN ",3, "ALERT"
	Set oExec=WshShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	FileSelected = oExec.StdOut.ReadLine
	if not FileSelected="" then 
		If FObj.FileExists(FileSelected) Then 
			If 	FObj.GetExtensionName(FileSelected)="docx" Or FObj.GetExtensionName(FileSelected)="doc" Then 
				If  OpenMSWordPaste(FileSelected) Then 
				    If CopyContent(FObj.GetBaseName(FileSelected)) Then 
						Call ClearImageFolder()
					End If 
					'WordObj.Quit 
				End If  
			Else 
				WshShell.Popup "SELECTED FILE IS NOT A WORD DOCUMENT :  PLEASE CHOOSE A WORD DOCUMENT",2, "ALERT"
				Call SelectExistingFile()
			End If 
		Else 
			WshShell.Popup " UNABLE TO OPEN FILE : MAY BE IT IS CORRUPTED : PLEASE CREATE A NEW FILE ",5,"ALERT"	
			Call NewFile()
		End If 
	else 
		Call Action()	
	End If 
End Function 		


'Function to get the last file

Function GetRecentFileName(ByVal Path)
	Dim RecentFile,file,objFolder,var
	var=0
		Set RecentFile = Nothing
		Set objFolder = FObj.GetFolder(Path)
  		For Each file in objFolder.Files
    			If RecentFile is Nothing Then
      				Set RecentFile = file
    			ElseIf file.DateLastModified > RecentFile.DateLastModified Then
      				Set RecentFile = file
    			End If
  		Next
  		If Not RecentFile is  Nothing  Then
  			var=Mid (RecentFile.Name,1,InStr (1,RecentFile.Name,".")-1)
  		End If
  		GetRecentFileName= var
End Function 

'Check whether the folder is empty or not 

Function IsVoid(Path)
Set FObj= CreateObject("Scripting.FileSystemObject")
	Dim X,f,Var,Obj
	X=0
	Set Obj=FObj.GetFolder(path)
	For Each f in  Obj.Files 
		if FObj.GetExtensionName(f)="png" or FObj.GetExtensionName(f)="jpg" or FObj.GetExtensionName(f)="jpeg" or FObj.GetExtensionName(f)="bmp"  then 
			X=X+1
		End if 
	Next
	If X="0" Then
		Var= true
	else
		var = False 
	End If
	IsVoid=Var
End Function 


Set WordObj= Nothing
Set FObj= Nothing