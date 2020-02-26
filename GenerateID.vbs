Dim Ip,c,Y,M,D,arr,WshShell,ExcelObj,FObj,ExcelPathgen,cou
Set WshShell = CreateObject("Wscript.Shell")
Set ExcelObj=CreateObject("Excel.Application")
Set FObj= CreateObject("Scripting.FileSystemObject")


Call Input()
Call GenerateID()

Function GenerateID()
 	Dim ID,SA,NAM,sum,var,dig
 	SA="0"
 	NAM="1"
 	If Not c=-1 Then 
		If InStr (UCase (Ip(1)),"M") And InStr (UCase (Ip(2)),"S") Then 
			ID=Y&M&D&Random("5000","9999")&SA&"8"
			gen="M"
			Cou="SA"
		ElseIf InStr (UCase (Ip(1)),"M") And InStr (UCase (Ip(2)),"N") Then
			ID=Y&M&D&Random("5000","9999")&NAM&"8" 
			gen="M"
			Cou="NAM"
		ElseIf InStr (UCase (Ip(1)),"F") And InStr (UCase (Ip(2)),"S") Then
			ID=Y&M&D&Random("0000","4999")&SA&"8"
			gen="F"
			Cou="SA"
		ElseIf InStr (UCase (Ip(1)),"F") And InStr (UCase (Ip(2)),"N") Then
			ID=Y&M&D&Random("0000","4999")&NAM&"8"
			gen="F"
			Cou="NAM"
		End If 
		
		sum=CalculateSum(ID,1,2)
		var=Concat(ID,2,2)*2
		sum=sum+CalculateSum(var,1,1)
		dig=CheckDigit(sum)
		dig=CheckDigit(10-dig)
		ID=ID&dig
	strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
	strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
		Set fso = CreateObject("Scripting.FileSystemObject")
    		If( fso.FileExists ("Generated National ID.txt")) Then 
    			Set f=fso.OpenTextFile("Generated National ID.txt",8)
    		Else
    			Set f=fso.CreateTextFile("Generated National ID.txt")
    		End If 
    		f.Write "DOB:"&D&"/"&M&"/"&Y& "  Gender:"&gen&"  Country:"&cou&	" NATIONAL ID: "& ID  &"  Generated at: " &Now()  &vbCrLf
   		f.Close
		wshshell.Popup "NATIONAL ID :"&ID&vbCrLf&"[NOTE: ID WILL BE AVAILABLE IN 'Generated National ID.txt' FILE",3,"OUTPUT"
	End If 
End Function 



Function Random(ByVal nLow, ByVal nHigh)
	Dim RN,l,z,i
	z=""
 	Randomize
  	RN=Int((nHigh - nLow + 1) * Rnd() + nLow)
  	l=Len(RN)
  	If Not l=4 Then 
  		For i=1 To 4-l Step 1
  			z=z&"0"
  		Next 
  		RN=z&RN
  	End If
  	Random=RN
End Function 
'Function to take Input and check input valiidation
Function Input()
	Dim var	
	var=InputBox ("ENTER DOB, GENDER, COUNTRY"&vbCrLf&" PLEASE FOLLOW THE EXACT FORMAT [DDMMYYYY M/F SA/NAM] ")
	var = Trim(var)
	Do While InStr(1, var, "  ")
  		var = Replace(var, "  ", " ")
	Loop
	Ip=Split(var," ")
	c=Ubound(Ip)
	If c=2 Then
		D=Mid(Ip(0),1,2)
		M=Mid(Ip(0),3,2)
		If Len(Ip(0))="8" Then 
			Y=Mid(Ip(0),7,2)
		ElseIf Len(Ip(0))="6" Then 
			Y=Mid(Ip(0),5,2)
		Else 
			wshshell.Popup "ERROR : WRONG DOB",1,"ERROR"
			Input()
		End If 
	ElseIf c=-1 Then 
		Exit Function 
	Else  
		wshshell.Popup "ERROR : WRONG NUMBER OF ARGUMENTS",1,"ERROR"
		Input()
	End If 
	If IsDate ( D&"/"&M&"/"&Y ) Then 
		If InStr (UCase (Ip(1)),"M") Or InStr (UCase (Ip(1)),"F") Then 
			If InStr (UCase (Ip(2)),"S") Or  InStr (UCase (Ip(2)),"N") Then
				 Exit Function
			Else 
				wshshell.Popup "ERROR : WRONG COUNTRY. ",1,"ERROR"
				Input()
			End If 
		Else
			wshshell.Popup "ERROR : WRONG GENDER. ",1,"ERROR"
			Input()
		End If 
	Else 
		wshshell.Popup "ERROR : WRONG DOB ",1,"ERROR"
		Input()
	End If 
End Function 
'Function to calculate sum 
Function CalculateSum(ByVal arr,ByVal StartingIndex, ByVal jump)
	Dim s
	s=0
	For i=StartingIndex To Len (arr) Step jump
		s=s+Mid(arr,i,1)	
	Next
	CalculateSum=s
End Function 
'Function to concatination// used for id validation
Function Concat(ByVal arr,ByVal StartingIndex, ByVal jump)
	Dim s
	s=""
	For i=StartingIndex To Len (arr) Step jump
		s=s&Mid(arr,i,1)	
	Next
	Concat=s
End Function

'Function to check how many digits and return string accordingly// used for id validation
Function CheckDigit(ByVal str)
	Dim var
	If Len (str)=2 Then 
		var=Mid(str,2,1)
	Else 
	 	var=str 
	End If
	CheckDigit=var 
End Function 	