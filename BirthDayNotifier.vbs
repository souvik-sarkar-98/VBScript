Option Explicit
Dim ExcelObj,objShell,objFSO,Book,sheetName,ExcelPath,TextPath,objSheet,rowCount,colCount,i,j,Var,c1,c2,ArrToday(),ArrTomm(),objFileToWrite

Set ExcelObj = CreateObject("Excel.Application")
Set objShell=CreateObject("WScript.Shell")
Set objFSO=CreateObject("Scripting.FileSystemObject")

ExcelPath = "C:\Users\USER\MyFolder\BirthdayList.xlsx"
TextPath=objShell.ExpandEnvironmentStrings("%userprofile%")&"\Birthday.txt"
sheetName = "BirthdayList"

c1=0
c2=0
ReDim PRESERVE ArrToday(1)
ReDim PRESERVE ArrTomm(1)

Set Book = ExcelObj.Workbooks.Open(ExcelPath)
Set objSheet = Book.Sheets(sheetName)
Set objFileToWrite =objFSO.OpenTextFile (TextPath,2,True)

rowCount = objSheet.UsedRange.Rows.Count
colCount = objSheet.UsedRange.Columns.Count 
	 if rowCount > 1 then 
		for i=1 to colCount step 1
			if InStr(UCase(objSheet.Cells(1,i)),"NAME") Then
				Var=i
			End If
			if InStr(UCase(objSheet.Cells(1,i)),"DATE") Then
				For j=2 to rowCount Step 1
					if Cdate(ConvertToDate(objSheet.Cells(j,i)))=Date() Then
						ReDim PRESERVE ArrToday(c1)
						ArrToday(c1)= Trim (UCase(objSheet.Cells(j,Var)))
						c1=c1+1
						msgbox("Today is "&objSheet.Cells(j,Var)&"'s BirthDay.Convey Best Wishes")
					end if 
					If DateDiff("d",Date(),CDate(ConvertToDate(objSheet.Cells(j,i))))="1" Then
						ReDim PRESERVE ArrTomm(c2)
						ArrTomm(c2)=Trim (UCase(objSheet.Cells(j,Var)))
						c2=c2+1
						msgbox("Tommorrow is "&objSheet.Cells(j,Var)&"'s BirthDay.Convey Best Wishes")
					end if 
				Next
			End If
		Next
		If UBound(ArrToday)>1 Or UBound(ArrTomm)>1 Then 
			call WriteInText(ArrToday,"Today is "," 's BirthDay.Convey"," Best Wishes")
			objFileToWrite.WriteLine("")
			call WriteInText(ArrTomm,"Tommorrow is "," 's BirthDay.Convey"," Best Wishes")
		End If 
	End if 	
ExcelObj.ActiveWorkbook.Close
Function WriteInText(byval arr(),ByVal arg1,ByVal arg2,ByVal arg3)
	Dim i,str,names
	str=""
	If UBound (arr)=0 Then 
		objFileToWrite.WriteLine(arg1&arr(0)&arg2&str&arg3)
	Else 
		names=arr(0)
		str=" Them"
		For i=1 To UBound(arr) Step 1 
			names=names&","&arr(i)
		Next 
		objFileToWrite.WriteLine(arg1&names&arg2&str&arg3)
	End If 
	str=""
End Function

Function ConvertToDate(ByVal Date)
	Dim str1,str2,monthList,i,j
	
	monthList=Array ("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","Oct","NOV","DEC")
	Date = Trim(Date)
	str1=Mid (Date,1,1)
	If IsNumeric (Mid (Date,2,1)) Then 
		str1=str1&Mid (Date,2,1)
	End If 
	For i=1 To Len (Date)-2 Step 1
		For j=0 To 11 Step 1 
			If (UCase (Mid (Date,i,3))=monthList(j)) Then 
				str2=monthList(j)
			End If 
		Next 
	Next 
	ConvertToDate=str1&"-"&str2
End Function 
Set objSheet =Nothing
set Book = Nothing
Set ExcelObj = Nothing